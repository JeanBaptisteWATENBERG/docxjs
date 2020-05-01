import { SectionProperties, HeaderAndFooterType } from "../dom/document";
import { ns, Columns, Column } from "../dom/common";
import * as xml from './common';

export function parseSectionProperties(elem: Element): SectionProperties {
    var section = <SectionProperties>{};

    for (let e of xml.elements(elem, ns.wordml)) {
        switch (e.localName) {
            case "pgSz":
                section.pageSize = {
                    width: xml.lengthAttr(e, ns.wordml, "w"),
                    height: xml.lengthAttr(e, ns.wordml, "h"),
                    orientation: xml.stringAttr(e, ns.wordml, "orient")
                }
                break;

            case "pgMar":
                section.pageMargins = {
                    left: xml.lengthAttr(e, ns.wordml, "left"),
                    right: xml.lengthAttr(e, ns.wordml, "right"),
                    top: xml.lengthAttr(e, ns.wordml, "top"),
                    bottom: xml.lengthAttr(e, ns.wordml, "bottom"),
                    header: xml.lengthAttr(e, ns.wordml, "header"),
                    footer: xml.lengthAttr(e, ns.wordml, "footer"),
                    gutter: xml.lengthAttr(e, ns.wordml, "gutter"),
                };
                break;

            case "cols":
                section.columns = parseColumns(e);
                break;

            case "footerReference":
                section.footers = {
                    ...section.footers,
                    [xml.stringAttr(e, ns.wordml, "type")]: {refId: xml.stringAttr(e, ns.relationships, "id")}
                }
                break;

            case "headerReference":
                section.headers = {
                    ...section.headers,
                    [xml.stringAttr(e, ns.wordml, "type")]: {refId: xml.stringAttr(e, ns.relationships, "id")}
                }
                break;
        }
    }

    return section;
}

function parseColumns(elem: Element): Columns {
    return {
        numberOfColumns: xml.intAttr(elem, ns.wordml, "num"),
        space: xml.lengthAttr(elem, ns.wordml, "space"),
        separator: xml.boolAttr(elem, ns.wordml, "sep"),
        equalWidth: xml.boolAttr(elem, ns.wordml, "equalWidth", true),
        columns: xml.elements(elem, ns.wordml, "col")
            .map(e => <Column>{
                width: xml.lengthAttr(e, ns.wordml, "w"),
                space: xml.lengthAttr(e, ns.wordml, "space")
            })
    };
}