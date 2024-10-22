import saveWebPart from "./saveWebPart";
const isoDate = new Date().toISOString();
describe("SaveWebParts", () => {
  const blobXml = `
  <WebPart xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://schemas.microsoft.com/WebPart/v2"
  ID="bec11dcb-1b31-4505-a82e-ff2837fb09fd">
  <Title>myContentEditor${isoDate}</Title>
  <FrameType>Default</FrameType>
  <Description>Allows authors to enter rich text content.</Description>
  <IsIncluded>true</IsIncluded>
  <ZoneID>Header</ZoneID>
  <PartOrder>0</PartOrder>
  <FrameState>Normal</FrameState>
  <Height />
  <Width />
  <AllowRemove>true</AllowRemove>
  <AllowZoneChange>true</AllowZoneChange>
  <AllowMinimize>true</AllowMinimize>
  <AllowConnect>true</AllowConnect>
  <AllowEdit>true</AllowEdit>
  <AllowHide>true</AllowHide>
  <IsVisible>true</IsVisible>
  <DetailLink />
  <HelpLink />
  <HelpMode>Modeless</HelpMode>
  <Dir>Default</Dir>
  <PartImageSmall />
  <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
  <PartImageLarge>/_layouts/images/mscontl.gif</PartImageLarge>
  <IsIncludedFilter />
  <ExportControlledProperties>true</ExportControlledProperties>
  <ConnectionID>00000000-0000-0000-0000-000000000000</ConnectionID>
  <ID>g_bec11dcb_1b31_4505_a82e_ff2837fb09fd</ID>
  <Assembly>Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
  <TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>
  <ContentLink xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor" />
  <Content xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor"><![CDATA[​​&lt;p&gt;here is the content for myContentEditor&lt;/p&gt;]]></Content>
  <PartStorage xmlns="http://schemas.microsoft.com/WebPart/v2/ContentEditor" />
</WebPart>
`;
  const blobStorageKey = "bec11dcb-1b31-4505-a82e-ff2837fb09fd";
  it("Save an XML blob", async () => {
    const res = await saveWebPart({
      pageURL:
        "http://objectpoint/sites/spws/operations/StaticPages/getWebPart.aspx",
      storageKey: blobStorageKey,
      webPartXml: blobXml,
      allowTypeChange: true,
      storage: "Shared",
    });
    expect(res.status === 200).toBeTruthy();
    expect(res.data.success).toBeTruthy();
  });
});
