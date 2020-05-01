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
    private parser: DocumentParser;

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
        d.parser = parser;

        return d.zip.loadAsync(blob).then(z => {
            return d.loadContentType().then(() => d);
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

    loadHeaderOrFooter(id: string): PromiseLike<DocumentElement> {
        return this.loadResource(this.docRelations, id, "text")
            .then(resource => resource ? this.parser.parseHeaderOrFooter(resource) : null);
    }

    getHyperlinkTarget(id: string): string  {
        const rel = this.docRelations.find(x => x.id == id);
        return rel.target;
    }

    private loadContentType() {
        const contentTypePart = this.zip.files['[Content_Types].xml'];
        if (!contentTypePart) {
            throw new Error("Invalid office open xml document, missing [Content_Types].xml");
        }
        
        return contentTypePart.async("text").then(xml => {
            const parts = this.parser.parseContentTypeFile(xml);

            const getRelPath = (path: string) => {
                if (!path) return path;
                const beginning = path.substr(0, path.lastIndexOf("/") + 1);
                const remaining = path.replace(beginning, "");
                return beginning + "_rels/" + remaining + ".rels";
            }

            const files = [
                this.loadPart(PartType.DocumentRelations, normalizePath(getRelPath(parts.get(PartType.Document)))),
                this.loadPart(PartType.FontRelations, normalizePath(getRelPath(parts.get(PartType.FontTable)))),
                this.loadPart(PartType.NumberingRelations, normalizePath(getRelPath(parts.get(PartType.Numbering)))),
                this.loadPart(PartType.Style, normalizePath(parts.get(PartType.Style))),
                this.loadPart(PartType.FontTable, normalizePath(parts.get(PartType.FontTable))),
                this.loadPart(PartType.Numbering, normalizePath(parts.get(PartType.Numbering))),
                this.loadPart(PartType.Document, normalizePath(parts.get(PartType.Document)))
            ];

            return Promise.all(files.filter(x => x != null));
        })
    }

    private loadResource(relations: IDomRelationship[], id: string, outputType: JSZip.OutputType = "base64") {
        let rel = relations.find(x => x.id == id);
        return rel ? this.zip.files[rel.target.startsWith("/") ? normalizePath(rel.target) : ("word/" + rel.target)].async(outputType) : Promise.resolve(null);
    }

    private loadPart(part: PartType, partPath: string) {
        var f = this.zip.files[partPath];

        return f ? f.async("text").then(xml => {
            switch (part) {
                case PartType.FontRelations:
                    this.fontRelations = this.parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.DocumentRelations:
                    this.docRelations = this.parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.NumberingRelations:
                    this.numRelations = this.parser.parseDocumentRelationsFile(xml);
                    break;

                case PartType.Style:
                    this.styles = this.parser.parseStylesFile(xml);
                    break;

                case PartType.Numbering:
                    this.numbering = this.parser.parseNumberingFile(xml);
                    break;

                case PartType.Document:
                    this.document = this.parser.parseDocumentFile(xml);
                    break;

                case PartType.FontTable:
                    this.fontTable = this.parser.parseFontTable(xml);
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