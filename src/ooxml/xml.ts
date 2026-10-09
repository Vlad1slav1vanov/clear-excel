const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";
const XML_ENTITIES = new Map([
  ["amp", "&"],
  ["quot", '"'],
  ["apos", "'"],
  ["lt", "<"],
  ["gt", ">"],
]);
const SPREADSHEET_NAMESPACES = new Set([
  "",
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://purl.oclc.org/ooxml/spreadsheetml/main",
]);
const XML_NAME = /^[A-Za-z_\u00c0-\uffff][A-Za-z0-9_.:\-\u00b7\u00c0-\uffff]*$/;

interface XmlAttribute {
  name: string;
  localName: string;
  namespaceUri: string;
  value: string;
  valueStart: number;
  valueEnd: number;
}

export interface XmlElement {
  name: string;
  localName: string;
  namespaceUri: string;
  attributes: XmlAttribute[];
  start: number;
  openingEnd: number;
  closingStart: number;
  end: number;
  parent: XmlElement | undefined;
}

interface XmlCollection {
  body: string;
  full: string;
  opening: string;
  element: XmlElement;
}

interface XmlEdit {
  start: number;
  end: number;
  text: string;
}

interface XmlFrame {
  element: XmlElement;
  namespaces: ReadonlyMap<string, string>;
}

export function getAttribute(xml: string, name: string): string | undefined {
  return openingAttributes(xml).find((attribute) => attribute.name === name)?.value;
}

function openingAttributes(xml: string): XmlAttribute[] {
  if (!xml.startsWith("<")) {
    return parseAttributes(` ${xml}`, -1);
  }
  const opening = xml.slice(1, findTagEnd(xml, 0) - 1);
  const name = opening.match(/^\S+?(?=[\s/]|$)/)?.[0] ?? "";
  return parseAttributes(opening.slice(name.length).replace(/\/\s*$/, ""), name.length + 1);
}

export function replaceAttribute(xml: string, name: string, value: string): string {
  const attribute = openingAttributes(xml).find((candidate) => candidate.name === name);
  return attribute === undefined
    ? xml
    : xml.slice(0, attribute.valueStart) + value + xml.slice(attribute.valueEnd);
}

export function getNamespacedAttribute(
  element: XmlElement,
  name: string,
  namespaceUris: readonly string[],
): string | undefined {
  return element.attributes.find(
    (attribute) => attribute.localName === name && namespaceUris.includes(attribute.namespaceUri),
  )?.value;
}

export function decodeXml(value: string): string {
  return value.replace(/&([^;]+);/g, (_entity, reference: string) => {
    const replacement = XML_ENTITIES.get(reference);
    if (replacement !== undefined) {
      return replacement;
    }
    const numeric = reference.startsWith("#x")
      ? /^#x[\da-f]+$/i.test(reference) && Number.parseInt(reference.slice(2), 16)
      : /^#\d+$/.test(reference) && Number.parseInt(reference.slice(1), 10);
    if (
      typeof numeric === "number" &&
      (numeric === 9 || numeric === 10 || numeric === 13 || numeric >= 32) &&
      numeric <= 0x10ffff &&
      !(numeric >= 0xd800 && numeric <= 0xdfff)
    ) {
      return String.fromCodePoint(numeric);
    }
    throw new Error(`Invalid XML character reference: &${reference};`);
  });
}

export function isSpreadsheetElement(element: XmlElement, name?: string): boolean {
  return (
    SPREADSHEET_NAMESPACES.has(element.namespaceUri) &&
    (name === undefined || element.localName === name)
  );
}

// The scanner keeps only the open-element stack, preserving source spans without
// materializing a DOM or interpreting text, CDATA, comments, or processing instructions.
function* scanXml(xml: string): Generator<{ element: XmlElement; opening: boolean }> {
  const stack: XmlFrame[] = [];
  let position = 0;
  while (position < xml.length) {
    const start = xml.indexOf("<", position);
    if (start === -1) {
      break;
    }
    if (
      xml.startsWith("<!--", start) ||
      xml.startsWith("<![CDATA[", start) ||
      xml.startsWith("<?", start)
    ) {
      const delimiter = xml.startsWith("<!--", start)
        ? "-->"
        : xml.startsWith("<?", start)
          ? "?>"
          : "]]>";
      const end = xml.indexOf(
        delimiter,
        start + (delimiter === "?>" ? 2 : delimiter === "-->" ? 4 : 9),
      );
      if (end === -1) {
        throw new Error("Unterminated XML comment, CDATA, or processing instruction");
      }
      position = end + delimiter.length;
      continue;
    }
    if (xml.startsWith("<!", start)) {
      throw new Error("Unsupported XML declaration");
    }
    const end = findTagEnd(xml, start);
    const body = xml.slice(start + 1, end - 1);
    if (body.startsWith("/")) {
      const name = body.slice(1).trim();
      const frame = stack.pop();
      if (!frame || frame.element.name !== name) {
        throw new Error(`Mismatched XML closing element: ${name}`);
      }
      frame.element.closingStart = start;
      frame.element.end = end;
      yield { element: frame.element, opening: false };
    } else {
      const name = body.match(/^\S+?(?=[\s/]|$)/)?.[0] ?? "";
      validateName(name);
      const selfClosing = /\/\s*$/.test(body);
      const attributes = parseAttributes(
        body.slice(name.length).replace(/\/\s*$/, ""),
        start + 1 + name.length,
      );
      const inherited = stack.at(-1)?.namespaces ?? new Map([["xml", XML_NAMESPACE]]);
      const declarations = attributes.filter(
        (attribute) => attribute.name === "xmlns" || attribute.name.startsWith("xmlns:"),
      );
      let namespaces = inherited;
      if (declarations.length > 0) {
        const updated = new Map(inherited);
        for (const attribute of declarations) {
          const prefix = attribute.name === "xmlns" ? "" : attribute.name.slice(6);
          if (
            prefix === "xmlns" ||
            (prefix === "xml") !== (attribute.value === XML_NAMESPACE) ||
            attribute.value === XMLNS_NAMESPACE
          ) {
            throw new Error(`Invalid XML namespace declaration: ${attribute.name}`);
          }
          updated.set(prefix, attribute.value);
        }
        namespaces = updated;
      }
      const attributeNames = new Set<string>();
      for (const attribute of attributes) {
        attribute.namespaceUri =
          attribute.name === "xmlns" || attribute.name.startsWith("xmlns:")
            ? XMLNS_NAMESPACE
            : resolveNamespace(attribute.name, namespaces, false);
        const expandedName = `${attribute.namespaceUri}\0${attribute.localName}`;
        if (attributeNames.has(expandedName)) {
          throw new Error(`Duplicate XML attribute: ${attribute.name}`);
        }
        attributeNames.add(expandedName);
      }
      const element: XmlElement = {
        name,
        localName: name.split(":").at(-1) ?? name,
        namespaceUri: resolveNamespace(name, namespaces, true),
        attributes,
        start,
        openingEnd: end,
        closingStart: end,
        end,
        parent: stack.at(-1)?.element,
      };
      yield { element, opening: true };
      if (selfClosing) {
        yield { element, opening: false };
      } else {
        stack.push({ element, namespaces });
      }
    }
    position = end;
  }
  if (stack.length > 0) {
    throw new Error(`Unclosed XML element: ${stack.at(-1)?.element.name}`);
  }
}

function findTagEnd(xml: string, start: number): number {
  let quote = "";
  for (let position = start + 1; position < xml.length; position += 1) {
    const character = xml[position];
    if (quote) {
      if (character === quote) {
        quote = "";
      }
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === ">") {
      return position + 1;
    } else if (character === "<") {
      throw new Error("Invalid XML opening element");
    }
  }
  throw new Error("Unterminated XML opening element");
}

function validateName(name: string): void {
  if (!XML_NAME.test(name) || name.split(":").length > 2 || name.endsWith(":")) {
    throw new Error(`Invalid XML name: ${name}`);
  }
}

function parseAttributes(xml: string, offset: number): XmlAttribute[] {
  const attributes: XmlAttribute[] = [];
  const names = new Set<string>();
  let position = 0;
  while (position < xml.length) {
    const match = /^\s+([^\s=]+)\s*=\s*(["'])/.exec(xml.slice(position));
    if (!match) {
      if (xml.slice(position).trim() !== "") {
        throw new Error("Invalid XML attributes");
      }
      break;
    }
    const name = match[1] ?? "";
    validateName(name);
    if (names.has(name)) {
      throw new Error(`Duplicate XML attribute: ${name}`);
    }
    names.add(name);
    const valueStart = position + match[0].length;
    const valueEnd = xml.indexOf(match[2] ?? '"', valueStart);
    if (valueEnd === -1 || xml.slice(valueStart, valueEnd).includes("<")) {
      throw new Error("Invalid XML attribute value");
    }
    attributes.push({
      name,
      localName: name.split(":").at(-1) ?? name,
      namespaceUri: "",
      value: decodeXml(xml.slice(valueStart, valueEnd)),
      valueStart: offset + valueStart,
      valueEnd: offset + valueEnd,
    });
    position = valueEnd + 1;
  }
  return attributes;
}

function resolveNamespace(
  name: string,
  namespaces: ReadonlyMap<string, string>,
  element: boolean,
): string {
  const colon = name.indexOf(":");
  if (colon === -1) {
    return element ? (namespaces.get("") ?? "") : "";
  }
  const prefix = name.slice(0, colon);
  const namespace = namespaces.get(prefix);
  if (!namespace) {
    throw new Error(`Undeclared XML namespace prefix: ${prefix}`);
  }
  return namespace;
}

export function findElements(xml: string, name: string): XmlElement[] {
  const elements: XmlElement[] = [];
  for (const { element, opening } of scanXml(xml)) {
    if (opening && element.localName === name) {
      elements.push(element);
    }
  }
  return elements;
}

export function mapOpeningElements(
  xml: string,
  transform: (element: XmlElement, opening: string) => string,
): string {
  const edits: XmlEdit[] = [];
  for (const { element, opening } of scanXml(xml)) {
    if (opening) {
      const original = xml.slice(element.start, element.openingEnd);
      const text = transform(element, original);
      if (text !== original) {
        edits.push({ start: element.start, end: element.openingEnd, text });
      }
    }
  }
  return applyEdits(xml, edits);
}

export function removeElements(xml: string, predicate: (element: XmlElement) => boolean): string {
  const edits: XmlEdit[] = [];
  for (const { element, opening } of scanXml(xml)) {
    if (!opening && predicate(element)) {
      edits.push({ start: element.start, end: element.end, text: "" });
    }
  }
  // Elements close inside-out; removing an outer element also removes its descendants.
  // oxlint-disable-next-line unicorn/no-array-sort
  edits.sort((left, right) => left.start - right.start || right.end - left.end);
  return applyEdits(xml, edits);
}

function applyEdits(xml: string, edits: readonly XmlEdit[]): string {
  const chunks: string[] = [];
  let position = 0;
  for (const edit of edits) {
    if (edit.start < position) {
      continue;
    }
    chunks.push(xml.slice(position, edit.start), edit.text);
    position = edit.end;
  }
  chunks.push(xml.slice(position));
  return chunks.join("");
}

export function findCollection(xml: string, name: string): XmlCollection | undefined {
  const element = findElements(xml, name).find((candidate) => isSpreadsheetElement(candidate));
  return element === undefined
    ? undefined
    : {
        body: xml.slice(element.openingEnd, element.closingStart),
        full: xml.slice(element.start, element.end),
        opening: xml.slice(element.start, element.openingEnd),
        element,
      };
}

export function replaceCollection(xml: string, name: string, body: string, count: number): string {
  const collection = findCollection(xml, name);
  if (!collection) {
    return xml;
  }
  const attribute = collection.element.attributes.find((candidate) => candidate.name === "count");
  const opening = (
    attribute
      ? collection.opening.slice(0, attribute.valueStart - collection.element.start) +
        count +
        collection.opening.slice(attribute.valueEnd - collection.element.start)
      : collection.opening.replace(/\/?\s*>$/, ` count="${count}">`)
  ).replace(/\/\s*>$/, ">");
  return (
    xml.slice(0, collection.element.start) +
    opening +
    body +
    `</${collection.element.name}>` +
    xml.slice(collection.element.end)
  );
}
