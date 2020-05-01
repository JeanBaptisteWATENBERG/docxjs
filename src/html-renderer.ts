import { Document } from './document';
import { IDomStyle, IDomTable, IDomStyleValues, IDomNumbering, OpenXmlElement } from './dom/dom';
import { Length } from './dom/common';
import { Options } from './docx-preview';
import { DocumentElement, SectionProperties, HeadersOrFooters, HeaderAndFooterType } from './dom/document';
import { ContainerBase, ElementBase } from './elements/element-base';
import { Break } from './elements/break';
import { Paragraph } from './elements/paragraph';
import { Run } from './elements/run';
import { Table } from './elements/table';
import { RenderContext } from './dom/render-context';

export class HtmlRenderer {

    inWrapper: boolean = true;
    className: string = "docx";
    document: Document;
    options: Options;

    private _renderContext: RenderContext = new RenderContext();

    constructor(public htmlDocument: HTMLDocument) {
        this._renderContext.html = htmlDocument;
    }

    render(document: Document, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
        this.document = document;
        this.options = options;

        this._renderContext.options = options;
        this._renderContext.className = this.className;
        this._renderContext.document = document;

        styleContainer = styleContainer || bodyContainer;

        removeAllElements(styleContainer);
        removeAllElements(bodyContainer);

        appendComment(styleContainer, "docxjs library predefined styles");
        styleContainer.appendChild(this.renderDefaultStyle());
        appendComment(styleContainer, "docx document styles");
        styleContainer.appendChild(this.renderStyles(document.styles));

        if (document.numbering) {
            appendComment(styleContainer, "docx document numbering styles");
            styleContainer.appendChild(this.renderNumbering(document.numbering, styleContainer));
        }

        if(!options.ignoreFonts)
            this.renderFontTable(document.fontTable, styleContainer);

        var wrapper = this.renderWrapper();
        
        bodyContainer.appendChild(wrapper);
        const container = this.inWrapper ? wrapper : bodyContainer
        this.renderSections(container, document.document);
    }

    renderFontTable(fonts: any[], styleContainer: HTMLElement) {
        // try to resolve external fonts
        for(let f of fonts.filter(x => !x.refId)) {
            appendComment(styleContainer, `Importing Google Font ${f.name}`);
            styleContainer.appendChild(createGoogleFontElement(f.name))
        }
        // load embedded fonts
        for(let f of fonts.filter(x => x.refId)) {
            this.document.loadFont(f.refId, f.fontKey).then(fontData => {
                var cssTest = `@font-face {
                    font-family: "${f.name}";
                    src: url(${fontData});
                }`;

                appendComment(styleContainer, `Font ${f.name}`);
                styleContainer.appendChild(createStyleElement(cssTest));
            });
        }
    }

    processClassName(className: string) {
        if (!className)
            return this.className;

        return `${this.className}_${className}`;
    }

    processStyles(styles: IDomStyle[]) {
        var stylesMap: Record<string, IDomStyle> = {};

        for (let style of styles.filter(x => x.id != null)) {
            stylesMap[style.id] = style;
        }

        for (let style of styles.filter(x => x.basedOn)) {
            var baseStyle = stylesMap[style.basedOn];

            if (baseStyle) {
                for (let styleValues of style.styles) {
                    var baseValues = baseStyle.styles.filter(x => x.target == styleValues.target);

                    if (baseValues && baseValues.length > 0)
                        this.copyStyleProperties(baseValues[0].values, styleValues.values);
                }
            }
            else if (this.options.debug)
                console.warn(`Can't find base style ${style.basedOn}`);
        }

        for (let style of styles) {
            style.id = this.processClassName(style.id);
        }

        return stylesMap;
    }

    processElement(element: OpenXmlElement) {
        if (element.children) {
            for (var e of element.children) {
                e.className = this.processClassName(e.className);
                e.parent = element;

                if (e instanceof Table) {
                    this.processTable(e);
                }
                else {
                    this.processElement(e);
                }
            }
        }
    }

    processTable(table: IDomTable) {
        for (var r of table.children) {
            for (var c of r.children) {
                c.style = this.copyStyleProperties(table.cellStyle, c.style, [
                    "border-left", "border-right", "border-top", "border-bottom",
                    "padding-left", "padding-right", "padding-top", "padding-bottom"
                ]);

                this.processElement(c);
            }
        }
    }

    copyStyleProperties(input: IDomStyleValues, output: IDomStyleValues, attrs: string[] = null): IDomStyleValues {
        if (!input)
            return output;

        if (output == null) output = {};
        if (attrs == null) attrs = Object.getOwnPropertyNames(input);

        for (var key of attrs) {
            if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
                output[key] = input[key];
        }

        return output;
    }

    createSection(className: string, props: SectionProperties, header?: DocumentElement, footer?: DocumentElement) {
        var elem = this.htmlDocument.createElement("section");
        
        elem.className = className;

        if (props) {
            if (props.pageMargins) {
                elem.style.paddingLeft = this.renderLength(props.pageMargins.left);
                elem.style.paddingRight = this.renderLength(props.pageMargins.right);
                elem.style.paddingTop = header ? this.renderLength(props.pageMargins.header) : this.renderLength(props.pageMargins.top);
                elem.style.paddingBottom = footer ? this.renderLength(props.pageMargins.footer) : this.renderLength(props.pageMargins.bottom);
            }

            if (props.pageSize) {
                if (!this.options.ignoreWidth)
                    elem.style.width = this.renderLength(props.pageSize.width);
                if (!this.options.ignoreHeight)
                    elem.style.height = this.renderLength(props.pageSize.height);
            }

            if (props.columns && props.columns.numberOfColumns) {
                elem.style.columnCount = `${props.columns.numberOfColumns}`;
                elem.style.columnGap = this.renderLength(props.columns.space);

                if (props.columns.separator) {
                    elem.style.columnRule = "1px solid black";
                }
            }
        }

        return elem;
    }

    async renderSections(into: HTMLElement, document: DocumentElement): Promise<HTMLElement[]> {
        var result = [];

        this.processElement(document);

        const sections = this.splitBySection(document.children);
        let sectionNumber = 1;

        while (sections.length > 0) {
            this._renderContext.currentPageNumber = sectionNumber;
            const section = sections.shift();
            const sectionProps = section.sectProps || document.props;
            const resolvedHeaderDefinitions = await Promise.all(Object.values(sectionProps.headers).map(({refId}) => this.document.loadHeaderOrFooter(refId)));
            const resolvedFooterDefinitions = await Promise.all(Object.values(sectionProps.footers).map(({refId}) => this.document.loadHeaderOrFooter(refId)));

            const toTypeIndex = (resolvedDefinitions: DocumentElement[]) => (type: string, index: number): { type: string; definition: DocumentElement; } => ({ type, definition: resolvedDefinitions[index] });
            const groupByType = (byType: {}, current: { type: string; definition: DocumentElement; }): {} => ({ ...byType, [current.type]: current.definition });
        
            const headersByType: {[type in HeaderAndFooterType]: DocumentElement} | {} = Object.keys(sectionProps.headers).map(toTypeIndex(resolvedHeaderDefinitions)).reduce(groupByType, {});
            const footersByType: {[type in HeaderAndFooterType]: DocumentElement} | {} = Object.keys(sectionProps.headers).map(toTypeIndex(resolvedFooterDefinitions)).reduce(groupByType, {});

            const pickedHeader = this.pickHeaderOrFooter(headersByType, sectionNumber);
            const pickedFooter = this.pickHeaderOrFooter(footersByType, sectionNumber);

            const sectionElement = this.createSection(this.className, sectionProps, pickedHeader, pickedFooter);

            into.appendChild(sectionElement);

            if (pickedHeader) {
                const header = this.htmlDocument.createElement("header");
                this.renderElements(pickedHeader.children, header);
                sectionElement.appendChild(header);
            }

            const main = this.htmlDocument.createElement("main");
            sectionElement.appendChild(main);

            if (pickedFooter) {
                const footer = this.htmlDocument.createElement("footer");
                this.renderElements(pickedFooter.children, footer);
                sectionElement.appendChild(footer);
            }

            const {remainingElementsAfterConstraintReached} = this.renderElements(section.elements, main, true);
            if (remainingElementsAfterConstraintReached && remainingElementsAfterConstraintReached.length > 0) {
                if (sections.length > 0) {
                    sections[0].elements.unshift(...remainingElementsAfterConstraintReached);
                } else {
                    const newSection = { sectProps: sectionProps, elements: remainingElementsAfterConstraintReached };
                    sections.push(newSection);
                }
            }

            result.push(sectionElement);
            sectionNumber++;
        }

        this.htmlDocument.querySelectorAll('.total-pages').forEach((elem: HTMLElement) => elem.innerText = `${sectionNumber - 1}`)

        return result;
    }

    private pickHeaderOrFooter(headersOrFootersByType: {[type in HeaderAndFooterType]: DocumentElement} | {}, sectionNumber: number): DocumentElement | undefined {
        if (headersOrFootersByType['first'] && sectionNumber === 1) {
            return headersOrFootersByType['first'];
        }
        else if (headersOrFootersByType['even'] && sectionNumber % 2 === 0) {
            return headersOrFootersByType['even'];
        }
        else if (headersOrFootersByType['default']) {
            return headersOrFootersByType['default'];
        }
        return undefined;
    }

    splitBySection(elements: OpenXmlElement[]): { sectProps: SectionProperties, elements: OpenXmlElement[] }[] {
        var current = { sectProps: null, elements: [] };
        var result = [current];

        function splitElement(elem: ContainerBase, revert: boolean) {
            var children = elem.children;
            var newElem = Object.create(Object.getPrototypeOf(elem));
            Object.assign(newElem, elem);

            let [f,s] = revert ? [elem, newElem] : [newElem, elem];

            f.children = children.slice(pBreakIndex);
            s.children = children.slice(0, rBreakIndex)

            return newElem;
        }

        for(let elem of elements) {
            current.elements.push(elem);

            if(elem instanceof Paragraph)
            {
                var sectProps = elem.props.sectionProps;
                var pBreakIndex = -1;
                var rBreakIndex = -1;
                
                if(this.options.breakPages && elem.children) {
                    pBreakIndex = elem.children.findIndex((r: Run) => {
                        rBreakIndex = r.children?.findIndex(t => (t instanceof Break) && t.break == "page") ?? -1;
                        return rBreakIndex != -1;
                    });
                }
    
                if(sectProps || pBreakIndex != -1) {
                    current.sectProps = sectProps;
                    current = { sectProps: null, elements: [] };
                    result.push(current);
                }

                if(pBreakIndex != -1) {
                    let breakRun = elem.children[pBreakIndex] as Run;
                    let splitRun = rBreakIndex < breakRun.children.length - 1;

                    if(pBreakIndex < elem.children.length - 1 || splitRun) {
                        current.elements.push(splitElement(elem, false));

                        if(splitRun) {
                            elem.children.push(splitElement(breakRun, true));
                        }
                    }
                }
            }
        }

        return result;
    }

    renderLength(l: Length): string {
        return !l ? null : `${l.value}${l.type}`;
    }

    renderWrapper() {
        var wrapper = document.createElement("div");

        wrapper.className = `${this.className}-wrapper`

        return wrapper;
    }

    renderDefaultStyle() {
        var styleText = `.${this.className}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
                .${this.className}-wrapper section.${this.className} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
                .${this.className} { color: black; }
                section.${this.className} { box-sizing: border-box; }
                .${this.className} table { border-collapse: collapse; }
                .${this.className} table td, .${this.className} table th { vertical-align: top; }
                .${this.className} p { margin: 0pt; }
                .${this.className} p:empty:before { content: ' '; white-space: pre; }
                
                section.${this.className} {
                    display: flex;
                    flex-flow: column;
                    height: 100%;
                  }
                  
                  section.${this.className} header {
                    flex: 0 1 auto;
                  }
                  
                  section.${this.className} main {
                    flex: 1 1 auto;
                  }
                  
                  section.${this.className} footer {
                    flex: 0 1 auto;
                  }
                
                `;

        return createStyleElement(styleText);
    }

    renderNumbering(styles: IDomNumbering[], styleContainer: HTMLElement) {
        var styleText = "";
        var rootCounters = [];

        for (var num of styles) {
            var selector = `p.${this.numberingClass(num.id, num.level)}`;
            var listStyleType = "none";

            if (num.levelText && num.format == "decimal") {
                let counter = this.numberingCounter(num.id, num.level);

                if (num.level > 0) {
                    styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
                        "counter-reset": counter
                    });
                }
                else {
                    rootCounters.push(counter);
                }

                styleText += this.styleToString(`${selector}:before`, {
                    "content": this.levelTextToContent(num.levelText, num.id),
                    "counter-increment": counter
                });

                styleText += this.styleToString(selector, {
                    "display": "list-item",
                    "list-style-position": "inside",
                    "list-style-type": "none",
                    ...num.style
                });
            }
            else if (num.bullet) {
                let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

                styleText += this.styleToString(`${selector}:before`, {
                    "content": "' '",
                    "display": "inline-block",
                    "background": `var(${valiable})`
                }, num.bullet.style);

                this.document.loadNumberingImage(num.bullet.src).then(data => {
                    var text = `.${this.className}-wrapper { ${valiable}: url(${data}) }`;
                    styleContainer.appendChild(createStyleElement(text));
                });
            }
            else {
                listStyleType = this.numFormatToCssValue(num.format);
            }

            styleText += this.styleToString(selector, {
                "display": "list-item",
                "list-style-position": "inside",
                "list-style-type": listStyleType,
                ...num.style
            });
        }

        if (rootCounters.length > 0) {
            styleText += this.styleToString(`.${this.className}-wrapper`, {
                "counter-reset": rootCounters.join(" ")
            });
        }

        return createStyleElement(styleText);
    }

    renderStyles(styles: IDomStyle[]): HTMLElement {
        var styleText = "";
        var stylesMap = this.processStyles(styles);

        for (let style of styles) {
            var subStyles =  style.styles;

            if(style.linked) {
                var linkedStyle = style.linked && stylesMap[style.linked];

                if (linkedStyle)
                    subStyles = subStyles.concat(linkedStyle.styles);
                else if(this.options.debug)
                    console.warn(`Can't find linked style ${style.linked}`);
            }

            for (var subStyle of subStyles) {
                var selector = "";

                if (style.target == subStyle.target)
                    selector += `${style.target}.${style.id}`;
                else if (style.target)
                    selector += `${style.target}.${style.id} ${subStyle.target}`;
                else
                    selector += `.${style.id} ${subStyle.target}`;

                if (style.isDefault && style.target)
                    selector = `.${this.className} ${style.target}, ` + selector;

                styleText += this.styleToString(selector, subStyle.values);
            }
        }

        return createStyleElement(styleText);
    }

    renderElements(elems: OpenXmlElement[], into?: HTMLElement, heightConstrained?: boolean): {renderedElements: Node[], remainingElementsAfterConstraintReached: OpenXmlElement[]} {
        if(elems == null)
            return null;

        var result = elems.map((e: ElementBase) => ({originalElement: e, renderedElement: e.render(this._renderContext)})).filter(e => e.renderedElement != null);

        if(into) {
            const appendedElements = [];
            const remainingElements = [...result];
            for(let c of result) {
                const containerBeforeHeight = into.getBoundingClientRect().height;
                into.appendChild(c.renderedElement);
                const containerAfterHeight = into.getBoundingClientRect().height;
                const containsPageBreak = (c.renderedElement as HTMLElement).querySelector('.page-break');
                
                if (containsPageBreak) {
                    appendedElements.push(c.renderedElement);
                    remainingElements.shift();
                    return {renderedElements: appendedElements, remainingElementsAfterConstraintReached: remainingElements.map(e => e.originalElement) };
                } else if (heightConstrained && containerBeforeHeight !== containerAfterHeight) {
                    into.removeChild(c.renderedElement);
                    return {renderedElements: appendedElements, remainingElementsAfterConstraintReached: remainingElements.map(e => e.originalElement) };
                } else {
                    appendedElements.push(c.renderedElement);
                    remainingElements.shift();
                }
            }
        }

        return {renderedElements: result.map(e => e.renderedElement), remainingElementsAfterConstraintReached: [] };
    }

    numberingClass(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }

    styleToString(selectors: string, values: IDomStyleValues, cssText: string = null) {
        let result = selectors + " {\r\n";

        for (const key in values) {
            result += `  ${key}: ${values[key]};\r\n`;
        }

        if (cssText)
            result += ";" + cssText;

        return result + "}\r\n";
    }

    numberingCounter(id: string, lvl: number) {
        return `${this.className}-num-${id}-${lvl}`;
    }

    levelTextToContent(text: string, id: string) {
        var result = text.replace(/%\d*/g, s => {
            let lvl = parseInt(s.substring(1), 10) - 1;
            return `"counter(${this.numberingCounter(id, lvl)})"`;
        });

        return '"' + result + '"';
    }

    numFormatToCssValue(format: string) {
        var mapping = {
            "none": "none",
            "bullet": "disc",
            "decimal": "decimal",
            "lowerLetter": "lower-alpha",
            "upperLetter": "upper-alpha",
            "lowerRoman": "lower-roman",
            "upperRoman": "upper-roman",
        };

        return mapping[format] || format;
    }
}

function appentElements(container: HTMLElement, children: HTMLElement[]) {
    for (let c of children)
        container.appendChild(c);
}

function removeAllElements(elem: HTMLElement) {
    while (elem.firstChild) {
        elem.removeChild(elem.firstChild);
    }
}

function createStyleElement(cssText: string) {
    var styleElement = document.createElement("style");
    styleElement.type = "text/css";
    styleElement.innerHTML = cssText;
    return styleElement;
}

function createGoogleFontElement(fontName: string) {
    /// Import font from google font using <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Tangerine">
    const fontLinkElement = document.createElement("link");
    fontLinkElement.rel = "stylesheet";
    fontLinkElement.href = `https://fonts.googleapis.com/css?family=${fontName}`
    return fontLinkElement;
}

function appendComment(elem: HTMLElement, comment: string) {
    elem.appendChild(document.createComment(comment));
}