import { defaults, Response } from "..";
import ResponseError from "./responseError";

interface RequestOptions {
  webService: string;
  webURL?: string;
}

class Request {
  private envelope = ``;
  private webService: string;
  private webURL?: string;
  xhr: XMLHttpRequest;

  constructor({ webURL, webService }: RequestOptions) {
    this.webService = webService;
    this.webURL = webURL;

    // Create XHR
    let xhr = new XMLHttpRequest();
    xhr.open("POST", `${this.webURL}/_vti_bin/${this.webService}.asmx`);
    xhr.setRequestHeader("Content-Type", `text/xml; charset="utf-8"`);
    this.xhr = xhr;
  }

  /**
   * Adds the body XML between the header and footer.
   *
   * @param body The body XML
   * @example
   * // Creates an envelope for a getList request
   * createEnvelope(`<GetList xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>Announcements</listName></GetList>`)
   */
  createEnvelope = (body: string) => {
    this.envelope = `
    <soap:Envelope
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
      xmlns:xsd="http://www.w3.org/2001/XMLSchema"
      xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        <soap:Body>
          ${body}
        </soap:Body>
    </soap:Envelope>`;

    return this;
  };

  send = (): Promise<Response> =>
    new Promise((resolve, reject) => {
      // TODO: Add Error Handling
      this.xhr.onreadystatechange = () => {
        if (this.xhr.readyState === 4) {
          if (this.xhr.status === 200) {
            resolve({
              responseText: this.xhr.responseText,
              responseXML: this.xhr.responseXML!,
              status: this.xhr.status,
              statusText: this.xhr.statusText,
            });
          } else {
            // Create response error
            const error = new ResponseError(this.xhr);

            reject(error);
          }
        }
      };

      // Send Request
      this.xhr.send(this.envelope);
    });
}

export default Request;
