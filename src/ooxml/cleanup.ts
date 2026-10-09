import { DRAWING_ELEMENTS } from "./constants.js";
import { findElements, isSpreadsheetElement, removeElements } from "./xml.js";

const THREADED_COMMENT_NAMESPACES = new Set([
  "",
  "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
  "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
]);

export function removeWorksheetDrawingMarkup(xml: string): string {
  return removeElements(
    xml,
    (element) => isSpreadsheetElement(element) && DRAWING_ELEMENTS.includes(element.localName),
  );
}

export function removeThreadedCommentExtensions(xml: string): string {
  const extensions = new Set<number>();
  for (const name of ["threadedComments", "threadedComment"]) {
    for (const element of findElements(xml, name)) {
      if (!THREADED_COMMENT_NAMESPACES.has(element.namespaceUri)) {
        continue;
      }
      let ancestor = element.parent;
      while (ancestor) {
        if (isSpreadsheetElement(ancestor, "ext")) {
          extensions.add(ancestor.start);
          break;
        }
        ancestor = ancestor.parent;
      }
    }
  }
  return removeElements(xml, (element) => extensions.has(element.start));
}
