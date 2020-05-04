import { ContainerBase, renderStyleValues } from "./element-base";
import { ParagraphProperties } from "../dom/paragraph";
import { RenderContext } from "../dom/render-context";
import { appendClass } from "../utils";
import { element } from "../parser/xml-serialize";
import { IDomStyleValues } from "../dom/dom";
import { Run } from "./run";

@element("p")
export class Paragraph extends ContainerBase {
    defaultRunStyle: IDomStyleValues = {};
    props: ParagraphProperties = {} as ParagraphProperties;

    render(ctx: RenderContext): Node {
        var elem = this.renderContainer(ctx, "p");

        if (this.props.numbering) {
            var numberingClass = ctx.numberingClass(this.props.numbering.id, this.props.numbering.level);
            elem.className = appendClass(elem.className, numberingClass);
        }

        if (this.children.length > 0) {
            if (this.children.find(c => c instanceof Run && (!c.style || Object.keys(c.style).length === 0))) {
                this.style = {...this.style, ...this.defaultRunStyle};
                renderStyleValues(this.style, elem);
            }
        } else {
            this.style = {...this.style, ...this.defaultRunStyle};
            renderStyleValues(this.style, elem);
        }

        return elem;
    }
}