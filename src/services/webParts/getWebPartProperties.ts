import { defaults } from "../..";
import { SpwsRequest, SpwsError } from "../../classes";
import { WebServices } from "../../enum";
import { SpwsResponse, WebPartProperties } from "../../types";

type GetWebPartProps = {
  pageURL: string;
  webURL?: string;
};

export interface Operation extends SpwsResponse {
  data: WebPartProperties[];
}

/**
 * @description Fetches the Web Part properties for a specified page within a SharePoint site.
 * This function retrieves all the Web Parts on a given Web Part page, including their metadata,
 * such as IDs, titles, visibility status, and layout details, as well as the full XML representation of each Web Part.
 *
 * @param {string} pageURL - The relative URL of the Web Part page (e.g., "/sites/mysite/pages/sample.aspx")
 * where the Web Part resides.
 * @param {string} [webURL] - The optional base URL of the SharePoint site. If not provided, it defaults to the value in `defaults.webURL`.
 * @returns {Promise<Operation>} - A promise that resolves to an Operation object containing an array of Web Part properties.
 */
const getWebPartProperties = async ({
  pageURL,
  webURL,
}: GetWebPartProps): Promise<Operation> => {
  try {
    // Create request object
    const req = new SpwsRequest({
      webService: WebServices.WebPartPages,
      webURL: webURL || defaults.webURL,
    });

    // Create envelope for SOAP request
    req.createEnvelope(
      `<GetWebPartProperties xmlns="http://microsoft.com/sharepoint/webpartpages">
        <pageUrl>${pageURL}</pageUrl>
        <behavior>Version3</behavior>
        <storage>Shared</storage>
      </GetWebPartProperties>`
    );

    // Send request
    const res = await req.send();
    const responseXML = res.responseXML;

    if (!responseXML) {
      throw new SpwsError({ message: "Invalid response XML" });
    }

    // Get all WebPart nodes from the response
    const allNodes = responseXML.getElementsByTagName("WebPart");

    if (allNodes.length === 0) {
      return {
        ...res,
        data: [],
      };
    }
    // Function to get text content of child node or attribute
    const getChildNodeText = (parent: Element, tagName: string) => {
      const child = parent.querySelector(tagName);
      return child ? child.textContent || "" : "";
    };

    // Parse each WebPart node and map to WebPartProperties
    const webPartProperties: WebPartProperties[] = Array.from(allNodes).map(
      (node: Element) => {
        return {
          webPartXml: new XMLSerializer().serializeToString(node), // The full XML for the WebPart

          // Attributes (these are still fetched as attributes)
          xmlns_xsi: node.getAttribute("xmlns:xsi") || "",
          xmlns_xsd: node.getAttribute("xmlns:xsd") || "",
          xmlns: node.getAttribute("xmlns") || "",
          ID: node.getAttribute("ID") || "",

          // Child nodes (these are fetched as XML elements)
          Title: getChildNodeText(node, "Title"),
          FrameType: getChildNodeText(node, "FrameType"),
          Description: getChildNodeText(node, "Description"),
          IsIncluded:
            getChildNodeText(node, "IsIncluded").toUpperCase() === "TRUE",
          ZoneID: getChildNodeText(node, "ZoneID"),
          PartOrder: parseInt(getChildNodeText(node, "PartOrder") || "0"),
          FrameState: getChildNodeText(node, "FrameState"),
          Height: getChildNodeText(node, "Height") || undefined,
          Width: getChildNodeText(node, "Width") || undefined,
          AllowRemove:
            getChildNodeText(node, "AllowRemove").toUpperCase() === "TRUE",
          AllowZoneChange:
            getChildNodeText(node, "AllowZoneChange").toUpperCase() === "TRUE",
          AllowMinimize:
            getChildNodeText(node, "AllowMinimize").toUpperCase() === "TRUE",
          AllowConnect:
            getChildNodeText(node, "AllowConnect").toUpperCase() === "TRUE",
          AllowEdit:
            getChildNodeText(node, "AllowEdit").toUpperCase() === "TRUE",
          AllowHide:
            getChildNodeText(node, "AllowHide").toUpperCase() === "TRUE",
          IsVisible:
            getChildNodeText(node, "IsVisible").toUpperCase() === "TRUE",
          DetailLink: getChildNodeText(node, "DetailLink") || undefined,
          HelpLink: getChildNodeText(node, "HelpLink") || undefined,
          HelpMode: getChildNodeText(node, "HelpMode") || "",
          Dir: getChildNodeText(node, "Dir") || "",
          PartImageSmall: getChildNodeText(node, "PartImageSmall") || undefined,
          MissingAssembly:
            getChildNodeText(node, "MissingAssembly") || undefined,
          PartImageLarge: getChildNodeText(node, "PartImageLarge") || undefined,
          IsIncludedFilter:
            getChildNodeText(node, "IsIncludedFilter") || undefined,
          ExportControlledProperties:
            getChildNodeText(
              node,
              "ExportControlledProperties"
            ).toUpperCase() === "TRUE",
          ConnectionID: getChildNodeText(node, "ConnectionID") || "",
          ID_: getChildNodeText(node, "ID") || "",
          Assembly: getChildNodeText(node, "Assembly") || "",
          TypeName: getChildNodeText(node, "TypeName") || "",
          HeaderTitle: getChildNodeText(node, "HeaderTitle") || undefined,
        };
      }
    );

    // Return the array of WebPartProperties
    return {
      ...res,
      data: webPartProperties,
    };
  } catch (error: any) {
    throw new SpwsError(error);
  }
};

export default getWebPartProperties;
