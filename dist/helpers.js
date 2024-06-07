"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getByNotation = exports.downloadBlob = void 0;
function downloadURL(url, fileName) {
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();
}
function downloadBlob(data, fileName, mimeType) {
    const blob = new Blob([data], { type: mimeType });
    const url = window.URL.createObjectURL(blob);
    downloadURL(url, fileName);
    setTimeout(() => { return window.URL.revokeObjectURL(url); }, 100);
}
exports.downloadBlob = downloadBlob;
/**
 * Retrieves nested properties from an object using an array of accessors.
 */
function getDeep(obj, accessors) {
    let length = accessors.length;
    for (let i = 0; i < length; i++) {
        if (typeof obj === 'undefined')
            return undefined;
        if (Array.isArray(obj) && typeof accessors[i] === 'string')
            return obj.map(el => getDeep(el, accessors.slice(i)));
        obj = obj[accessors[i]];
    }
    return obj;
}
/**
 * Retrieves nested properties from an object using dot and bracket notation.
 */
const getByNotation = (obj, props) => {
    // Regular expression for extracting fields with deep notation data
    const notationRegex = /(?<=\[(?<qoute>['"]))[^'"\]].*?(?=\k<qoute>\])|(?<=\[)[^'"\]].*?(?=\])|[^.\["''\]]+(?=\.|\[|$)/g;
    const accessors = props.match(notationRegex);
    if (!accessors || !accessors.length)
        return obj;
    const accss = accessors.map(el => /^\d+$/.test(el) ? parseInt(el) : el);
    const value = getDeep(obj, accss);
    return value || '';
};
exports.getByNotation = getByNotation;
