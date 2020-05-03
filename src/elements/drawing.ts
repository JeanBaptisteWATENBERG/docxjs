import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

export class Drawing extends ContainerBase {

    posX = null;
    posY = null;
    style = {};

    render(ctx: RenderContext): Node {
        var elem = this.renderContainer(ctx, "div");

        elem.style.display = "inline-block";
        elem.style.textIndent = "0px";

        return elem 
    }
}