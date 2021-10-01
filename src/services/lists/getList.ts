// Constants
import { defaults, DefaultParameters, Response } from "../..";

// Enum
import ListAttributesEnum from "../../enum/listAttributes";
import FieldEnum from "../../enum/field";
import WebServices from "../../enum/webServices";

// Types
import FieldType from "../../types/field";
import List from "../../types/list";
import ListAttributes from "../../types/listAttributes";

// Classes
import Request from "../../classes/request";
export { default as ResponseError } from "../../classes/responseError";

// Utils
import escapeXml from "../../utils/escapeXml";

export interface GetListParameters extends DefaultParameters {
  /**
   * A string that contains either the title (not static name) or the GUID for the list.
   */
  listName: string;
  /**
   *  An array of attributes that are returned in the data object.
   *  Only available when parsing is true.
   *  If no attributes are supplied, all list attributes will be returned
   * */
  attributes?: ListAttributes[];
}

export interface GetListResponse extends Response {
  /**
   * The data object is available for any requests where parsed is true or an error occurs
   */
  data?: List;
}

/**
 * Returns a schema for the specified list.
 * @link https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-services/ms772709(v=office.12)
 * @example
 * ```
 * // Get list using default parameters
 * const list = await getList({ listName: "Announcements" });
 * // Get list on another site without parsing XML
 * const list = await getList({ listName: "Announcements", webURL: "/sites/hr", parse: false });
 * // Get list with only the Title and Fields parsed
 * const list = await getList({ listName: "Title", attributes: ["Title", "Fields"] })
 * ```
 */
const getList = ({
  listName,
  parse = defaults.parse,
  webURL = defaults.webURL,
  attributes = [],
}: GetListParameters): Promise<GetListResponse> => {
  return new Promise(async (resolve, reject) => {
    {
      // Create request object
      const req = new Request({ webService: WebServices.Lists, webURL });

      // Create envelope
      req.createEnvelope(
        `<GetList xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>${escapeXml(
          listName
        )}</listName></GetList>`
      );

      try {
        // Return request
        let res: GetListResponse = await req.send();

        // If parse is true
        if (parse) {
          // Get list from the responseXML
          const list: Element = res.responseXML.querySelector("List")!;

          // Create array of attributes either from params or all of the list attributes
          let attributesArray =
            attributes.length > 0
              ? attributes
              : Array.from(list.attributes).map((el) => el.name);

          // Create data object with only specified attributes
          res.data = attributesArray.reduce((object: List, attribute) => {
            object[attribute] = list.getAttribute(attribute) || "";
            return object;
          }, {});

          // If the attributes param is empty, or it included fields
          if (
            attributes.length === 0 ||
            attributes.includes(ListAttributesEnum.Fields)
          )
            // Add fields to data
            res.data[ListAttributesEnum.Fields] = Array.from(
              // Field attributes must be an array
              list.querySelectorAll(`${ListAttributesEnum.Fields} > Field`)
            ).map((fieldElement) => {
              // Create field object
              let field: FieldType = {};

              // If the field type is a choice field
              if (fieldElement.getAttribute(FieldEnum.Type) === "Choice") {
                // Add choicess to the field
                field.Choices = Array.from(
                  fieldElement.querySelectorAll("CHOICE")
                )
                  // Return text content
                  .map(({ textContent }) => textContent!)
                  // Remove empty choices
                  .filter((choice) => choice);
              }

              // Reduce field from available attributes
              return Array.from(fieldElement.attributes).reduce(
                (object: FieldType, element) => {
                  // Get field name and value
                  const key = element.nodeName;
                  let value: string | boolean = element.textContent || "";

                  // If the value is true or false
                  if (["TRUE", "FALSE"].includes(value)) {
                    // Cast to boolean
                    value = value === "TRUE";
                  }

                  // Assign key and prop
                  object[key] = value;
                  return object;
                },
                field
              );
            });
        }

        resolve(res);
      } catch (error: unknown) {
        reject(error);
      }
    }
  });
};

export default getList;
