import { OpenXmlElement } from "../dom/dom";
import { RenderContext } from "../dom/render-context";
import { appendClass } from "../utils";
import { sumSizeString } from "../parser/common";

export abstract class ElementBase implements OpenXmlElement {
    type: any;
    parent: OpenXmlElement;
    style: any;

    render(ctx: RenderContext): Node {
        return null;
    }
}

export abstract class ContainerBase extends ElementBase {
    children: ElementBase[] = [];
    childrenStyle: any = {};
    className: string;

    protected renderContainer<K extends keyof HTMLElementTagNameMap>(ctx: RenderContext, tagName: K): HTMLElementTagNameMap[K] {
        var elem = ctx.html.createElement(tagName);

        renderStyleValues(this.style, elem);

        if (this.className)
            elem.className = appendClass(elem.className, this.className);
        
        for(let n of this.children.map(c => {
            if ((c['posX'] && c['posX'].relative !== 'page') || (c['posY'] && c['posY'].relative !== 'page')) {
                elem.style.position = "relative";
            }
            if (this.childrenStyle) {
                c.style = {...this.childrenStyle, ...c.style};
            }
            return c.render(ctx);
        }).filter(x => x != null))
            elem.appendChild(n);

        return elem;
    }
}

export function renderStyleValues(style: any, ouput: HTMLElement) {
    if (style == null)
        return;

    for (let key in style) {
        if (style.hasOwnProperty(key)) {
            switch(key) {
                case 'bdr-top':
                case 'bdr-bottom':
                case 'bdr-left':
                case 'bdr-right':
                    break;
                case 'padding-top':
                case 'padding-bottom':
                case 'padding-left':
                case 'padding-right':
                    ouput.style[key] = sumSizeString(ouput.style[key], style[key]);
                    break;
                case 'bdr-top-space':
                    ouput.style["padding-top"] = sumSizeString(ouput.style["padding-top"], style[key]);
                    break;
                case 'bdr-bottom-space':
                    ouput.style["padding-bottom"] = sumSizeString(ouput.style["padding-bottom"], style[key]);
                    break;
                case 'bdr-left-space':
                    ouput.style["padding-left"] = sumSizeString(ouput.style["padding-left"], style[key]);
                    break;
                case 'bdr-right-space':
                    ouput.style["padding-right"] = sumSizeString(ouput.style["padding-right"], style[key]);
                    break;
                default:
                    ouput.style[key] = style[key];
            }
        }
    }
}