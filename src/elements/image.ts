import { ElementBase, renderStyleValues } from "./element-base"
import { RenderContext } from "../dom/render-context";
import { parseSizeFromSizeString } from "../parser/common";

export interface Box {
    left: number;
    right: number;
    top: number;
    bottom: number;
}

export class Image extends ElementBase {

    src: string;
    style: any = {};
    crop?: Box;
    stretch?: Box;

    render(ctx: RenderContext): Node {
        const result = ctx.html.createElement("img");
        console.log(this)

        if (ctx.document) {
            ctx.document.loadDocumentImage(this.src).then(x => {
                result.src = x;
            });
        }

        if (this.crop) {
            const cropWrapper = ctx.html.createElement("div");
            cropWrapper.appendChild(result);
            renderStyleValues(this.style, cropWrapper);
            cropWrapper.style.overflow = "hidden";
            const width = parseSizeFromSizeString(cropWrapper.style.width);
            const height = parseSizeFromSizeString(cropWrapper.style.height);
            const imageWidth = width / (1 - (this.crop.left/100_000) - (this.crop.right/100_000));
            const imageHeight = height / (1 - (this.crop.top / 100000) - (this.crop.bottom / 100000));
            
            result.style.marginLeft = `-${imageWidth * this.crop.left / 100_000}pt`;
            result.style.marginRight = `-${imageWidth * this.crop.right / 100_000}pt`;
            result.style.marginTop = `-${imageHeight * this.crop.top / 100_000}pt`;
            result.style.marginBottom = `-${imageHeight * this.crop.bottom / 100_000}pt`;
            
            result.style.width = `${imageWidth}pt`;
            result.style.height = `${imageHeight}pt`;
            return cropWrapper;
        } else {
            renderStyleValues(this.style, result);
        }

        return result;
    }
}