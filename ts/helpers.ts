function downloadURL(url: string, fileName: string): void {
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  a.style.display = 'none'
  document.body.appendChild(a)
  a.click()
  a.remove()
}

export function downloadBlob(data: Blob, fileName: string, mimeType: string): void {
  const blob = new Blob([data], { type: mimeType })
  const url = window.URL.createObjectURL(blob)
  downloadURL(url, fileName)
  setTimeout(() => { return window.URL.revokeObjectURL(url) }, 100)
}

/**
 * Retrieves nested properties from an object using an array of accessors.
 */
function getDeep(obj: Object, accessors: any[]): Object | undefined {
  let length = accessors.length
  for (let i = 0; i < length; i++) {
    if (typeof obj === 'undefined') return undefined
    if (Array.isArray(obj) && typeof accessors[i] === 'string') return obj.map(el => getDeep(el, accessors.slice(i)))
    obj = obj[accessors[i]]
  }
  return obj
}

/**
 * Retrieves nested properties from an object using dot and bracket notation.
 */
export const getByNotation = (obj: Object, props: string) => {
  // Regular expression for extracting fields with deep notation data
  const notationRegex = /(?<=\[(?<qoute>['"]))[^'"\]].*?(?=\k<qoute>\])|(?<=\[)[^'"\]].*?(?=\])|[^.\["''\]]+(?=\.|\[|$)/g
  const accessors = props.match(notationRegex)
  if (!accessors || !accessors.length) return obj
  const accss = accessors.map(el => /^\d+$/.test(el) ? parseInt(el) : el)
  const value = getDeep(obj, accss)
  return value || ''
}

