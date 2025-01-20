import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { DOMImplementation, XMLSerializer } from 'xmldom';

const IS_BROWSER = typeof window !== "undefined";

export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
    renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
    useBase64URL: boolean;
    renderChanges: boolean;
    renderComments: boolean;
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    useBase64URL: false,
    renderChanges: false,
    renderComments: false
}

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const domImpl = new DOMImplementation();
    const doc = domImpl.createDocument(null, null, null);
    const renderer = new HtmlRenderer(doc);
    return await renderer.render(document, bodyContainer, styleContainer, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
    const doc = await parseAsync(data, userOptions);
    await renderDocument(doc, bodyContainer, styleContainer, userOptions);
    return doc;
}

export async function renderToHtmlAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<string> {
    if (IS_BROWSER) {
        const root = await renderToHtmlElementAsync(data, userOptions);
        return root.outerHTML;
    }

    const root = await renderToHtmlElementAsync(data, userOptions);
    const serializer = new XMLSerializer();
    return serializer.serializeToString(root);
}

export async function renderToHtmlElementAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<HTMLElement> {
    if (IS_BROWSER) {
        const root = document.createElement('div');
        await renderAsync(data, root, undefined, userOptions);
        return root;
    }

    const domImpl = new DOMImplementation();
    const doc = domImpl.createDocument(null, null, null);
    const root = doc.createElement('div');
    await renderAsync(data, root, undefined, userOptions);
    handleXmlNode(root);
    return root;
}

function kebabCase(property: string) {
    return property.replace(/[A-Z]/g, char => `-${char.toLowerCase()}`);
}

function handleXmlNode(node: Node, depth = 0) {
    if (isElement(node)) {
        if (node.className) {
            node.setAttribute("class", node.className);
            node.className = null;
        }
        if ((node as any)._classList) {
            node.setAttribute("class", (node.classList as unknown as string[]).join(" "));
            (node as any)._classList = null;
        }
        if (node.innerHTML) {
            const textNode = node.ownerDocument.createTextNode(node.innerHTML);
            node.appendChild(textNode);
            node.innerHTML = null;
        }
        if ((node as any)._style) {
            var styles = Object.keys(node.style).map(key => {
                return `${kebabCase(key)}: ${node.style[key]}`;
            });
            node.setAttribute("style", styles.join("; "));
            (node as any)._style = null;
        }
        if (isTableCell(node)) {
            if (node.rowSpan) {
                node.setAttribute("rowspan", node.rowSpan.toString());
            }
            if (node.colSpan) {
                node.setAttribute("colspan", node.colSpan.toString());
            }
        }
        else if (isImage(node) && node.src) {
            node.setAttribute("src", node.src);
            node.src = null;
        }
    }

    if (node.childNodes) {
        for (let i = 0; i < node.childNodes.length; i++) {
            handleXmlNode(node.childNodes[i], depth + 1);
        }
    }
}

function isElement(value: Node): value is HTMLElement {
    return value.nodeType === 1; // Node.ELEMENT_NODE
}

function isTableCell(value: Node): value is HTMLTableCellElement {
    return (value as HTMLElement).tagName == "td";
}

function isImage(value: Node): value is HTMLImageElement {
    return (value as HTMLElement).tagName == "img";
}

function initXmlDom() {
    if (IS_BROWSER) {
        return;
    }

    const doc = new DOMImplementation().createDocument(null, null, null);
    const proto = Object.getPrototypeOf(Object.getPrototypeOf(doc));

    if (!proto.hasOwnProperty("firstElementChild")) {
        Object.defineProperty(proto, "firstElementChild", {
            get() {
                return getFirstElementChild(this);
            }
        });
    }

    if (!proto.hasOwnProperty("style")) {
        Object.defineProperty(proto, "style", {
            get() {
                return this._style || (this._style = {});
            },
            set(value) {
                this._style = value;
            }
        });
    }

    if (!proto.hasOwnProperty("classList")) {
        Object.defineProperty(proto, "classList", {
            get() {
                return this._classList || (this._classList = []);
            },
            set(value) {
                this._classList = value;
            }
        });
    }
}

function getFirstElementChild(node: Node) {
    if (!node) {
        return null;
    }
    const children = node.childNodes;
    for (let i = 0; i < children.length; i++) {
        if (children[i].nodeType === 1) { // Node.ELEMENT_NODE
            return children[i];
        }
    }
    return null;
}

initXmlDom();