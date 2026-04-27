import { readFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

export const MACRO_ENABLED_EXTENSION = "xlsm";
export const MACRO_ENABLED_CONTENT_TYPE =
  "application/vnd.ms-excel.sheet.macroEnabled.12";

const VBA_PROJECT_PATH = path.join(process.cwd(), "lib", "assets", "vbaProject.bin");
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const RELS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const DRAWING_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const VML_DRAWING_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
const CTRL_PROP_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
const VBA_PROJECT_REL_TYPE =
  "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
const COMMENTS_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

let vbaProjectBufferPromise;

function getVbaProjectBuffer() {
  if (!vbaProjectBufferPromise) {
    vbaProjectBufferPromise = readFile(VBA_PROJECT_PATH);
  }

  return vbaProjectBufferPromise;
}

function getNextRelationshipId(relationshipsXml) {
  const ids = [...relationshipsXml.matchAll(/Id="rId(\d+)"/g)].map((match) =>
    Number(match[1])
  );

  return `rId${Math.max(0, ...ids) + 1}`;
}

function createRelationshipsXml() {
  return `${XML_HEADER}\n<Relationships xmlns="${RELS_NAMESPACE}"></Relationships>`;
}

function escapeXml(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function createCommentsXml(notes) {
  const commentElements = notes
    .map(({ cellRef, text }) =>
      `<comment ref="${cellRef}" authorId="0"><text><r><t xml:space="preserve">${escapeXml(text)}</t></r></text></comment>`
    )
    .join("");
  return `${XML_HEADER}\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author></author></authors><commentList>${commentElements}</commentList></comments>`;
}


function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function addRelationship(relationshipsXml, type, target) {
  const existingRelationship = relationshipsXml.match(
    new RegExp(
      `<Relationship\\b(?=[^>]*Type="${escapeRegExp(type)}")(?=[^>]*Target="${escapeRegExp(target)}")[^>]*Id="([^"]+)"[^>]*/?>`
    )
  );

  if (existingRelationship) {
    return {
      id: existingRelationship[1],
      xml: relationshipsXml
    };
  }

  const id = getNextRelationshipId(relationshipsXml);
  const relationship = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;

  return {
    id,
    xml: relationshipsXml.replace("</Relationships>", `${relationship}</Relationships>`)
  };
}

function ensureDefaultContentType(contentTypesXml, extension, contentType) {
  if (contentTypesXml.includes(`Extension="${extension}"`)) {
    return contentTypesXml;
  }

  const defaultNode = `<Default Extension="${extension}" ContentType="${contentType}"/>`;
  return contentTypesXml.replace("</Types>", `${defaultNode}</Types>`);
}

function ensureOverrideContentType(contentTypesXml, partName, contentType) {
  const overridePattern = new RegExp(
    `<Override\\b(?=[^>]*PartName="${escapeRegExp(partName)}")[^>]*/>`
  );
  const overrideNode = `<Override PartName="${partName}" ContentType="${contentType}"/>`;

  if (overridePattern.test(contentTypesXml)) {
    return contentTypesXml.replace(overridePattern, overrideNode);
  }

  return contentTypesXml.replace("</Types>", `${overrideNode}</Types>`);
}

function ensureRootAttribute(xml, tagName, attributeName, attributeValue) {
  const openingTag = xml.match(new RegExp(`<${tagName}\\b[^>]*>`))?.[0];

  if (!openingTag || openingTag.includes(`${attributeName}=`)) {
    return xml;
  }

  const updatedOpeningTag = openingTag.replace(
    ">",
    ` ${attributeName}="${attributeValue}">`
  );

  return xml.replace(openingTag, updatedOpeningTag);
}

function ensureWorkbookCodeName(workbookXml) {
  if (/<workbookPr\b/.test(workbookXml)) {
    return workbookXml.replace(/<workbookPr\b([^>]*)>/, (match, attributes) => {
      if (attributes.includes("codeName=")) {
        return match;
      }

      return `<workbookPr codeName="ThisWorkbook"${attributes}>`;
    });
  }

  return workbookXml.replace(
    /(<workbook\b[^>]*>)/,
    '$1<workbookPr codeName="ThisWorkbook"/>'
  );
}

function ensureSheetCodeName(worksheetXml) {
  if (/<sheetPr\b/.test(worksheetXml)) {
    return worksheetXml.replace(/<sheetPr\b([^>]*)>/, (match, attributes) => {
      if (attributes.includes("codeName=")) {
        return match;
      }

      return `<sheetPr codeName="Sheet1"${attributes}>`;
    });
  }

  return worksheetXml.replace(
    /(<worksheet\b[^>]*>)/,
    '$1<sheetPr codeName="Sheet1"/>'
  );
}

function ensureWorksheetNamespaces(worksheetXml) {
  let updatedXml = worksheetXml;

  updatedXml = ensureRootAttribute(
    updatedXml,
    "worksheet",
    "xmlns:xdr",
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  );
  updatedXml = ensureRootAttribute(
    updatedXml,
    "worksheet",
    "xmlns:x14",
    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  );
  updatedXml = ensureRootAttribute(
    updatedXml,
    "worksheet",
    "xmlns:mc",
    "http://schemas.openxmlformats.org/markup-compatibility/2006"
  );

  return updatedXml;
}

function createDrawingXml() {
  return `${XML_HEADER}
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
    <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
      <xdr:twoCellAnchor>
        <xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp macro="" textlink="">
          <xdr:nvSpPr>
            <xdr:cNvPr id="1025" name="Button 1" hidden="1">
              <a:extLst>
                <a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1025"/></a:ext>
              </a:extLst>
            </xdr:cNvPr>
            <xdr:cNvSpPr/>
          </xdr:nvSpPr>
          <xdr:spPr bwMode="auto">
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:noFill/>
            <a:ln w="9525"><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln>
          </xdr:spPr>
          <xdr:txBody>
            <a:bodyPr vertOverflow="clip" wrap="square" lIns="36576" tIns="36576" rIns="36576" bIns="36576" anchor="ctr" upright="1"/>
            <a:lstStyle/>
            <a:p>
              <a:pPr algn="ctr" rtl="0"><a:defRPr sz="1000"/></a:pPr>
              <a:r>
                <a:rPr lang="en-GB" sz="1100" b="0" i="0" u="none" strike="noStrike" baseline="0">
                  <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
                  <a:latin typeface="Aptos Narrow"/>
                </a:rPr>
                <a:t>Add 100 new empty rows</a:t>
              </a:r>
            </a:p>
          </xdr:txBody>
        </xdr:sp>
        <xdr:clientData fPrintsWithSheet="0"/>
      </xdr:twoCellAnchor>
    </mc:Choice>
    <mc:Fallback/>
  </mc:AlternateContent>
</xdr:wsDr>`;
}

function createNoteVmlShapes(notes) {
  if (!notes.length) return "";

  const noteShapeType = `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
 </v:shapetype>`;

  const shapes = notes
    .map(({ row, col }, index) => {
      const shapeId = `_x0000_s${1026 + index}`;
      const anchorLeft = col + 1;
      const anchorRight = col + 3;
      const anchorTop = Math.max(0, row - 3);
      const anchorBottom = Math.max(anchorTop + 3, row);

      return `<v:shape id="${shapeId}" type="#_x0000_t202" style='position:absolute;margin-left:${80 + index * 80}pt;margin-top:0pt;width:108pt;height:60pt;z-index:${2 + index};visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto">
  <v:fill color2="#ffffe1"/>
  <v:shadow on="t" color="black" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style='mso-direction-alt:auto'>
   <div style='text-align:left'></div>
  </v:textbox>
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>${anchorLeft}, 15, ${anchorTop}, 2, ${anchorRight}, 28, ${anchorBottom}, 14</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>${row}</x:Row>
   <x:Column>${col}</x:Column>
  </x:ClientData>
 </v:shape>`;
    })
    .join("");

  return noteShapeType + shapes;
}

function createVmlDrawingXml(notes = []) {
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"
  path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
  <o:lock v:ext="edit" shapetype="t"/>
 </v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t201" style='position:absolute;
  z-index:1;mso-wrap-style:tight' o:button="t" fillcolor="buttonFace [67]" o:insetmode="auto">
  <v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>
  <o:lock v:ext="edit" rotation="t"/>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
   <div style='text-align:center'><font face="Aptos Narrow" size="220"
   color="#000000">Add 100 new empty rows</font></div>
  </v:textbox>
  <x:ClientData ObjectType="Button">
   <x:Anchor>4, 0, 2, 0, 5, 0, 3, 0</x:Anchor>
   <x:PrintObject>False</x:PrintObject>
   <x:AutoFill>False</x:AutoFill>
   <x:FmlaMacro>Add100Rows</x:FmlaMacro>
   <x:TextHAlign>Center</x:TextHAlign>
   <x:TextVAlign>Center</x:TextVAlign>
  </x:ClientData>
 </v:shape>${createNoteVmlShapes(notes)}</xml>`;
}

function createControlPropertiesXml() {
  return `${XML_HEADER}
<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Button" lockText="1"/>`;
}

function createWorksheetControlXml({ drawingRelationshipId, vmlRelationshipId, ctrlRelationshipId }) {
  return `<drawing r:id="${drawingRelationshipId}"/><legacyDrawing r:id="${vmlRelationshipId}"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><control shapeId="1025" r:id="${ctrlRelationshipId}" name="Button 1"><controlPr defaultSize="0" print="0" autoFill="0" autoPict="0" macro="Add100Rows"><anchor moveWithCells="1" sizeWithCells="1"><from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls></mc:Choice></mc:AlternateContent>`;
}

function getNextPartPath(zip, prefix, extension) {
  let index = 1;

  while (zip.file(`${prefix}${index}${extension}`)) {
    index += 1;
  }

  return `${prefix}${index}${extension}`;
}

function insertWorksheetControl(worksheetXml, controlXml) {
  if (worksheetXml.includes('macro="Add100Rows"')) {
    return worksheetXml;
  }

  if (worksheetXml.includes("<tableParts")) {
    return worksheetXml.replace("<tableParts", `${controlXml}<tableParts`);
  }

  return worksheetXml.replace("</worksheet>", `${controlXml}</worksheet>`);
}

export async function addMacroButtonToWorkbookBuffer(workbookBuffer, options = {}) {
  const { notes = [] } = options;
  const zip = await JSZip.loadAsync(workbookBuffer);
  const vbaProjectBuffer = await getVbaProjectBuffer();
  const drawingPath = getNextPartPath(zip, "xl/drawings/drawing", ".xml");
  const vmlDrawingPath = getNextPartPath(zip, "xl/drawings/vmlDrawing", ".vml");
  const controlPropertiesPath = getNextPartPath(zip, "xl/ctrlProps/ctrlProp", ".xml");

  zip.file("xl/vbaProject.bin", vbaProjectBuffer);
  zip.file(drawingPath, createDrawingXml());
  zip.file(vmlDrawingPath, createVmlDrawingXml(notes));
  zip.file(controlPropertiesPath, createControlPropertiesXml());

  const workbookRelsPath = "xl/_rels/workbook.xml.rels";
  const workbookRelsXml = await zip.file(workbookRelsPath).async("string");
  const workbookVbaRelationship = addRelationship(
    workbookRelsXml,
    VBA_PROJECT_REL_TYPE,
    "vbaProject.bin"
  );
  zip.file(workbookRelsPath, workbookVbaRelationship.xml);

  const sheetRelsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const existingSheetRels = zip.file(sheetRelsPath)
    ? await zip.file(sheetRelsPath).async("string")
    : createRelationshipsXml();
  const drawingRelationship = addRelationship(
    existingSheetRels,
    DRAWING_REL_TYPE,
    `../${drawingPath.replace("xl/", "")}`
  );
  const vmlRelationship = addRelationship(
    drawingRelationship.xml,
    VML_DRAWING_REL_TYPE,
    `../${vmlDrawingPath.replace("xl/", "")}`
  );
  const controlRelationship = addRelationship(
    vmlRelationship.xml,
    CTRL_PROP_REL_TYPE,
    `../${controlPropertiesPath.replace("xl/", "")}`
  );
  let finalSheetRels = controlRelationship.xml;
  if (notes.length > 0) {
    const commentsRelationship = addRelationship(finalSheetRels, COMMENTS_REL_TYPE, "../comments1.xml");
    finalSheetRels = commentsRelationship.xml;
    zip.file("xl/comments1.xml", createCommentsXml(notes));
  }
  zip.file(sheetRelsPath, finalSheetRels);

  const workbookPath = "xl/workbook.xml";
  const workbookXml = await zip.file(workbookPath).async("string");
  zip.file(workbookPath, ensureWorkbookCodeName(workbookXml));

  const worksheetPath = "xl/worksheets/sheet1.xml";
  const worksheetXml = await zip.file(worksheetPath).async("string");
  const worksheetControlXml = createWorksheetControlXml({
    drawingRelationshipId: drawingRelationship.id,
    vmlRelationshipId: vmlRelationship.id,
    ctrlRelationshipId: controlRelationship.id
  });
  const updatedWorksheetXml = insertWorksheetControl(
    ensureWorksheetNamespaces(ensureSheetCodeName(worksheetXml)),
    worksheetControlXml
  );
  zip.file(worksheetPath, updatedWorksheetXml);

  const contentTypesPath = "[Content_Types].xml";
  let contentTypesXml = await zip.file(contentTypesPath).async("string");
  contentTypesXml = ensureDefaultContentType(
    contentTypesXml,
    "bin",
    "application/vnd.ms-office.vbaProject"
  );
  contentTypesXml = ensureDefaultContentType(
    contentTypesXml,
    "vml",
    "application/vnd.openxmlformats-officedocument.vmlDrawing"
  );
  contentTypesXml = ensureOverrideContentType(
    contentTypesXml,
    "/xl/workbook.xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
  );
  contentTypesXml = ensureOverrideContentType(
    contentTypesXml,
    `/${drawingPath}`,
    "application/vnd.openxmlformats-officedocument.drawing+xml"
  );
  contentTypesXml = ensureOverrideContentType(
    contentTypesXml,
    `/${controlPropertiesPath}`,
    "application/vnd.ms-excel.controlproperties+xml"
  );
  if (notes.length > 0) {
    contentTypesXml = ensureOverrideContentType(
      contentTypesXml,
      "/xl/comments1.xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
    );
  }
  zip.file(contentTypesPath, contentTypesXml);

  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE"
  });
}
