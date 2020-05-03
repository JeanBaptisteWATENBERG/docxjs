import { Options } from "../docx-preview";
import { Document } from "../document";

export class RenderContext {
    html: HTMLDocument;
    options: Options;
    className: string;
    document: Document;
    currentPageNumber: number = 1;
    lastRenderedElement?: Node;

    numberingClass(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }
}