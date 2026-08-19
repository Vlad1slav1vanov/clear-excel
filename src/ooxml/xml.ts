interface XmlCollection {
  body: string;
  full: string;
  opening: string;
}

export function getAttribute(xml: string, name: string): string | undefined {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = xml.match(new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*(["'])(.*?)\\1`, "i"));
  const value = match?.[2];
  return value === undefined ? undefined : decodeXml(value);
}

export function decodeXml(value: string): string {
  return value
    .replaceAll("&amp;", "&")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">");
}

export function findCollection(xml: string, name: string): XmlCollection | undefined {
  const match = xml.match(
    new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:\\w+:)?${name}\\s*>`, "i"),
  );
  if (!match) {
    return undefined;
  }
  const full = match[0];
  if (full === undefined) {
    return undefined;
  }
  return {
    body: match[1] ?? "",
    full,
    opening: full.slice(0, full.indexOf(">") + 1),
  };
}

export function replaceCollection(xml: string, name: string, body: string, count: number): string {
  const collection = findCollection(xml, name);
  if (!collection) {
    return xml;
  }
  const opening = /\bcount\s*=/.test(collection.opening)
    ? collection.opening.replace(/(\bcount\s*=\s*)(["'])\d+\2/i, `$1"${count}"`)
    : collection.opening.replace(/>$/, ` count="${count}">`);
  const prefix = collection.opening.match(/^<(\w+:)/)?.[1] ?? "";
  return xml.replace(collection.full, `${opening}${body}</${prefix}${name}>`);
}

export function matchElements(xml: string, name: string): string[] {
  return (
    xml.match(
      new RegExp(
        `<(?:\\w+:)?${name}\\b[^>]*\\/\\s*>|<(?:\\w+:)?${name}\\b[^>]*>[\\s\\S]*?<\\/(?:\\w+:)?${name}\\s*>`,
        "gi",
      ),
    ) ?? []
  );
}
