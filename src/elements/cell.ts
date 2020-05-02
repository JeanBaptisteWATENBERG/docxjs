import { ContainerBase } from "./element-base";
import { RenderContext } from "../dom/render-context";

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

        return elem;
    }
}

export type VMerge = 'continue' | 'restart' | 'end';
export interface CellProperties {
    gridSpan: number;
    vMerge: VMerge;
    rowSpan: number;
}