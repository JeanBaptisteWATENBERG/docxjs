import { ElementBase } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { fromAttribute } from "../parser/xml-serialize";

export type BreakType =  "page" | "lastRenderedPageBreak" | "textWrapping";

export class Break extends ElementBase {
    @fromAttribute("type")
    break: BreakType = "textWrapping";

    render(ctx: RenderContext) {
        if (this.break === 'page') {
            const pageBreak = ctx.html.createElement('span');
            pageBreak.className = 'page-break';
            return pageBreak;
        }
        return this.break == "textWrapping" ? ctx.html.createElement("br") : null;
    }
}
