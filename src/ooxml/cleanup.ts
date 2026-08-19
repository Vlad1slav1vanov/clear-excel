import { DRAWING_ELEMENTS } from "./constants.js";

export function removeWorksheetDrawingMarkup(xml: string): string {
  return DRAWING_ELEMENTS.reduce((result, name) => {
    const element = new RegExp(
      `<(?:\\w+:)?${name}\\b[^>]*(?:\\/\\s*>|>[\\s\\S]*?<\\/(?:\\w+:)?${name}\\s*>)`,
      "gi",
    );
    return result.replace(element, "");
  }, xml);
}

export function removeThreadedCommentExtensions(xml: string): string {
  return xml.replace(/<(?:\w+:)?ext\b[^>]*>[\s\S]*?<\/(?:\w+:)?ext\s*>/gi, (extension) =>
    /<(?:\w+:)?threadedComments?\b/i.test(extension) ? "" : extension,
  );
}
