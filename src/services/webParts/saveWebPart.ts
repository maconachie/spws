import { defaults } from "../..";
import { SpwsRequest, SpwsError } from "../../classes";
import { WebServices } from "../../enum";
import { SpwsResponse } from "../../types";
import { escapeXml } from "../../utils";

type SaveWebPartProps = {
  pageURL: string; // The relative URL of the Web Part page (e.g., "/sites/mysite/pages/sample.aspx").
  storage: "None" | "Personal" | "Shared"; // Defines the storage scope: None, Personal, or Shared.
  storageKey: string; // A unique key identifying the Web Part to be saved.
  webPartXml: string; // The XML string representing the Web Part configuration.
  allowTypeChange?: boolean; // Optional: whether to allow changing the Web Part type.
  webURL?: string; // Optional: Base URL of the SharePoint site. Defaults to `defaults.webURL` if not provided.
};

export interface Operation extends SpwsResponse {
  data: { success: boolean }; // The response includes a success flag indicating whether the save operation was successful.
}

/**
 * @description Saves the Web Part configuration for a specified Web Part on a SharePoint page.
 * This function sends the updated XML configuration for a given Web Part and applies it to the page.
 * It uses the `SaveWebPart` SOAP operation provided by the SharePoint WebPartPages web service.
 *
 * @param {string} pageURL - The relative URL of the page that contains the Web Part (e.g., "/sites/mysite/pages/sample.aspx").
 * @param {string} storageKey - A unique identifier (GUID) of the Web Part to be saved.
 * @param {string} storage - Specifies the storage scope: "None", "Personal", or "Shared". Defaults to "Shared".
 * @param {string} webPartXml - The XML representation of the Web Part configuration.
 * @param {string} [webURL] - Optional: The base URL of the SharePoint site. If not provided, defaults to the value in `defaults.webURL`.
 * @returns {Promise<Operation>} - A promise that resolves to an Operation object, including a success flag if the Web Part was saved successfully.
 */
const saveWebPart = async ({
  pageURL,
  storageKey,
  storage = "Shared",
  webPartXml,
  webURL,
}: SaveWebPartProps): Promise<Operation> => {
  try {
    // Create request object for the WebPartPages web service
    const req = new SpwsRequest({
      webService: WebServices.WebPartPages,
      soapAction: "http://microsoft.com/sharepoint/webpartpages/SaveWebPart",
      webURL: webURL || defaults.webURL,
    });

    // Create SOAP envelope for the SaveWebPart operation
    req.createEnvelope(
      `<SaveWebPart xmlns="http://microsoft.com/sharepoint/webpartpages">
        <pageUrl>${pageURL}</pageUrl>
        <storageKey>${storageKey}</storageKey>
        <webPartXml>${escapeXml(webPartXml)}</webPartXml>
        <storage>${storage}</storage>
      </SaveWebPart>`
    );

    // Send the SOAP request and await the response
    const res = await req.send();

    // Check for SOAP Fault errors in the response
    if (res.responseXML.querySelector("soap\\:Fault, Fault"))
      throw new SpwsError(res);

    // Return the result indicating the Web Part was successfully saved
    return {
      ...res,
      data: { success: true },
    };
  } catch (error: any) {
    throw new SpwsError(error);
  }
};

export default saveWebPart;
