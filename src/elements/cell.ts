import { ContainerBase, renderStyleValues } from "./element-base";
import { RenderContext } from "../dom/render-context";
import { xml } from "../document-parser";
import { parseSizeFromSizeString } from "../parser/common";

export class Cell extends ContainerBase {
    props: CellProperties = {} as CellProperties;
    
    render(ctx: RenderContext): Node {
        var elem = this.renderContainer(ctx, "td");

        if(this.props.gridSpan)
            elem.colSpan = this.props.gridSpan;

        if(this.props.rowSpan) {
            elem.rowSpan = this.props.rowSpan;
        }

        if(this.props.vMerge && this.props.vMerge !== 'restart') {
            return null;
        }

        if (!this.style["padding-left"]) {
            this.style["padding-left"] = xml.convertSize("115");
        }

        if (!this.style["padding-right"]) {
            this.style["padding-right"] = xml.convertSize("115");
        }

        if(this.style["width"]) {
            this.style["width"] = (parseSizeFromSizeString(this.style["width"]) - 
                                parseSizeFromSizeString(this.style["padding-right"]) -
                                parseSizeFromSizeString(this.style["padding-left"])) + "pt";
        }

        renderStyleValues(this.style, elem);

        return elem;
    }
}

export type VMerge = 'continue' | 'restart';
export interface CellProperties {
    gridSpan: number;
    vMerge: VMerge;
    rowSpan: number;
}