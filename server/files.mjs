import {DOMParser} from '@xmldom/xmldom';
import {ApiError} from './google.mjs';
export const MAX_IMAGE_BYTES=2*1024*1024;
const elements=new Set('svg g path rect circle ellipse line polyline polygon text tspan defs linearGradient radialGradient stop clipPath mask title desc use symbol'.split(' '));
const attributes=new Set('id x y x1 x2 y1 y2 cx cy r rx ry width height viewBox preserveAspectRatio d points transform fill fill-opacity fill-rule stroke stroke-width stroke-opacity stroke-linecap stroke-linejoin stroke-dasharray stroke-dashoffset opacity font-family font-size font-weight text-anchor dominant-baseline offset stop-color stop-opacity gradientUnits gradientTransform spreadMethod clip-path clipPathUnits mask maskUnits maskContentUnits href'.split(' '));
export function validateSvg(text) {
  if(/<!DOCTYPE|<!ENTITY|<\?(?!xml\s)|<\!\[CDATA/i.test(text))throw new ApiError('SVG contains unsupported active content');
  let broken=false;const doc=new DOMParser({errorHandler:{warning(){broken=true;},error(){broken=true;},fatalError(){broken=true;}}}).parseFromString(text,'image/svg+xml');
  if(broken||doc.documentElement?.localName!=='svg')throw new ApiError('Invalid SVG image');
  function visit(node){
    if(node.nodeType===1){
      if(!elements.has(node.localName)||node.namespaceURI!=='http://www.w3.org/2000/svg')throw new ApiError('SVG contains unsupported elements');
      for(let i=0;i<node.attributes.length;i++){
        const a=node.attributes[i],name=a.localName,value=a.value;
        if(a.name==='xmlns'&&value==='http://www.w3.org/2000/svg'||a.name==='xmlns:xlink'&&value==='http://www.w3.org/1999/xlink')continue;
        if(!attributes.has(name)||a.namespaceURI&&a.namespaceURI!=='http://www.w3.org/1999/xlink')throw new ApiError('SVG contains unsupported attributes');
        if(name==='href'&&!/^#[\w.-]+$/.test(value))throw new ApiError('SVG external links are not allowed');
        if(/url\(/i.test(value)&&!/^url\(#[\w.-]+\)$/.test(value))throw new ApiError('SVG external resources are not allowed');
        if(/javascript:|data:|https?:|@import|expression\(|[\\<>]/i.test(value))throw new ApiError('SVG contains unsafe content');
      }
    }else if(![3,8,9,7].includes(node.nodeType))throw new ApiError('Invalid SVG content');
    for(let c=node.firstChild;c;c=c.nextSibling)visit(c);
  }visit(doc);return text;
}
export function validateImage(file) {
  if(!file||typeof file.base64!=='string'||typeof file.name!=='string'||file.name.length>180||file.base64.length>Math.ceil(MAX_IMAGE_BYTES/3)*4+4)throw new ApiError('Image is invalid or exceeds 2 MB');
  if(!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(file.base64))throw new ApiError('Invalid image encoding');
  const bytes=Buffer.from(file.base64,'base64');if(!bytes.length||bytes.length>MAX_IMAGE_BYTES)throw new ApiError('Image must be between 1 byte and 2 MB');
  const extension=file.name.split('.').pop().toLowerCase();let mime;
  if(extension==='png'&&bytes.subarray(0,8).equals(Buffer.from([137,80,78,71,13,10,26,10])))mime='image/png';
  if(['jpg','jpeg'].includes(extension)&&bytes[0]===255&&bytes[1]===216&&bytes[bytes.length-2]===255&&bytes[bytes.length-1]===217)mime='image/jpeg';
  if(extension==='webp'&&bytes.subarray(0,4).toString()==='RIFF'&&bytes.subarray(8,12).toString()==='WEBP')mime='image/webp';
  if(extension==='svg'){validateSvg(bytes.toString('utf8'));mime='image/svg+xml';}
  if(!mime)throw new ApiError('File contents must match PNG, JPG, JPEG, SVG, or WebP');
  return {name:file.name,base64:bytes.toString('base64'),mime,extension};
}
