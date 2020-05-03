import { ContainerBase } from "./element-base";
import { SectionProperties } from "../dom/document";
import { RenderContext } from "../dom/render-context";

export class Section extends ContainerBase {
    props: SectionProperties = {} as SectionProperties;

    render(ctx: RenderContext): Node {

        return null;
    }
}