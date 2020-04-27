import * as JSZip from 'jszip';

import { DocumentParser } from './document-parser';
import { IDomRelationship, IDomStyle, IDomNumbering } from './dom/dom';
import { Font } from './dom/common';
import { DocumentElement } from './dom/document';

export enum PartType {
    Document = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    Numbering = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    FontTable = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    Style = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    DocumentRelations = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.relationships+xml", // not an official content type (implementation detail)
    NumberingRelations = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering.relationships+xml", // not an official content type (implementation detail)
    FontRelations = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable.relationships+xml", // not an official content type (implementation detail)
}


const normalizePath = (path: string) => {
    return path ? path.replace(/^\/+/, '') : path;
}

export class Document {
    private zip: JSZip = new JSZip();

    docRelations: IDomRelationship[] = null;
    fontRelations: IDomRelationship[] = null;
    numRelations: IDomRelationship[] = null;

    styles: IDomStyle[] = null;
    fonts: Font[] = null;
    fontTable: any;
    numbering: IDomNumbering[] = null;
    document: DocumentElement = null;

    static load(blob, parser: DocumentParser): PromiseLike<Document> {
        var d = new Document();

        return d.zip.loadAsync(blob).then(z => {
            return d.loadContentType(parser).then(() => d);
        });
    }

    loadDocumentImage(id: string): PromiseLike<string> {
        return this.loadResource(this.docRelations, id, "blob")
            .then(x => x ? URL.createObjectURL(x) : null);
    }

    loadNumberingImage(id: string): PromiseLike<string> {
        return this.loadResource(this.numRelations, id, "blob")
            .then(x => x ? URL.createObjectURL(x) : null);
    }

    loadFont(id: string, key: string): PromiseLike<string> {
        return this.loadResource(this.fontRelations, id, "uint8array")
            .then(x => x ? URL.createObjectURL(new Blob([deobfuscate(x, key)])) : x);
    }

    private loadContentType(parser: DocumentParser) {
        const contentTypePart = this.zip.files['[Content_Types].xml'];
        if (!contentTypePart) {
            throw new Error("Invalid office open xml document, missing [Content_Types].xml");
        }
        
        return contentTypePart.async("text").then(xml => {
            const parts = parser.parseContentTypeFile(xml);

            const getRelPath = (path: string) => {
                if (!path) return path;
                const beginning = path.substr(0, path.lastIndexOf("/") + 1);
                const remaining = path.replace(beginning, "");
                return beginning + "_rels/" + remaining + ".rels";
            }

            console.log(normalizePath(getRelPath(parts.get(PartType.Document))), contentTypePart, this.zip.files)
            const files = [
                this.loadPart(PartType.DocumentRelations, normalizePath(getRelPath(parts.get(PartType.Document))), parser),
                this.loadPart(PartType.FontRelations, normalizePath(getRelPath(parts.get(PartType.FontTable))), parser),
                this.loadPart(PartType.NumberingRelations, normalizePath(getRelPath(parts.get(PartType.Numbering))), parser),
                this.loadPart(PartType.Style, normalizePath(parts.get(PartType.Style)), parser),
                this.loadPart(PartType.FontTable, normalizePath(parts.get(PartType.FontTable)), parser),
                this.loadPart(PartType.Numbering, normalizePath(parts.get(PartType.Numbering)), parser),
                this.loadPart(PartType.Document, normalizePath(parts.get(PartType.Document)), parser)
            ];

            return Promise.all(files.filter(x => x != null));
        })
    }

    private loadResource(relations: IDomRelationship[], id: string, outputType: JSZip.OutputType = "base64") {
        let rel = relations.find(x => x.id == id);
        return rel ? this.zip.files[rel.target.startsWith("/") ? normalizePath(rel.target) : ("word/" + rel.target)].async(outputType) : Promise.resolve(null);
    }

    private loadPart(part: PartType, partPath: string, parser: DocumentParser) {
        var f = this.zip.files[partPath];

        return f ? f.async("text").then(xml => {
            switch (part) {
                case PartType.FontRelations:
                    this.fontRelations = parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.DocumentRelations:
                    this.docRelations = parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.NumberingRelations:
                    this.numRelations = parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.Style:
                    this.styles = parser.parseStylesFile(xml);
                    break;

                case PartType.Numbering:
                    this.numbering = parser.parseNumberingFile(xml);
                    break;

                case PartType.Document:
                    this.document = parser.parseDocumentFile(xml);
                    break;

                case PartType.FontTable:
                    this.fontTable = parser.parseFontTable(xml);
                    break;
            }

            return this;
        }) : null;
    }
}

export function deobfuscate(data: Uint8Array, guidKey: string): Uint8Array {
    const len = 16;
    const trimmed = guidKey.replace(/{|}|-/g, "");
    const numbers = new Array(len);
    
    for(let i = 0; i < len; i ++)
        numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);

    for (let i = 0; i < 32; i++)
        data[i] = data[i] ^ numbers[i % len]

    return data;
}