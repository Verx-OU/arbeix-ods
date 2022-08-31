import { XMLBuilder, XMLParser } from "fast-xml-parser";
import JSZip from "jszip";
import { Document, Item } from "./node";

const XML_OPTIONS = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  parseTagValue: false,
  parseAttributeValue: false,
} as const;
const parser = new XMLParser(XML_OPTIONS);
const builder = new XMLBuilder(XML_OPTIONS);

const CONTENT_XML = "content.xml" as const;

export default async function copyAndModifyODS(
  source: Buffer,
  target: NodeJS.WritableStream,
  fn: (doc: Document) => void
) {
  const zip = await JSZip.loadAsync(source);
  const contentFile = zip.file(CONTENT_XML);
  if (contentFile === null) throw new Error("No content in ODS");
  const data = await contentFile.async("string");
  const obj = parser.parse(data) as Item[];
  const doc = new Document(obj);

  fn(doc);

  const result = builder.build(obj);
  zip.file(CONTENT_XML, result);

  zip
    .generateNodeStream({
      type: "nodebuffer",
      streamFiles: true,
      compression: "DEFLATE",
    })
    .pipe(target);
}
