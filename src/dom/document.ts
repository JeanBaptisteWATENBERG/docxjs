import { Length, Columns } from "./common";
import { OpenXmlElement } from "./dom";

export interface PageSize {
    width: Length, 
    height: Length, 
    orientation: "landscape" | string 
}

export interface PageMargins {
    top: Length;
    right: Length;
    bottom: Length;
    left: Length;
    header: Length;
    footer: Length;
    gutter: Length;
}

export type HeaderAndFooterType = 'default' | 'even' | 'first';

export type HeadersOrFooters = {
    [type in HeaderAndFooterType]?: {refId: string};
};

export interface SectionProperties {
    pageSize: PageSize,
    pageMargins: PageMargins,
    columns: Columns;
    headers: HeadersOrFooters;
    footers: HeadersOrFooters;
}

export interface DocumentElement extends OpenXmlElement {
    props: SectionProperties;
}