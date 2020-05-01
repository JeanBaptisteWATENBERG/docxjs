import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { fromAttribute } from "../parser/xml-serialize";

export class Hyperlink extends ContainerBase {
    @fromAttribute("anchor")
    anchor: string;

    @fromAttribute("id")
    refId: string;

    render(ctx: RenderContext): Node {
        var a = this.renderContainer(ctx, "a");

        if(this.anchor)
            a.href = `#${this.anchor}`;

        if (this.refId) {
            a.href = ctx.document.getHyperlinkTarget(this.refId);
            a.target = '_blank';
            a.rel = "noopener noreferrer";
        }
        
        return a;
    }
}