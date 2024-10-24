type WebPartProperties = {
  // A string containing the XML for a Web Part to include the type and assembly name fields.
  webPartXml: string;
  xmlns_xsi: string; // Namespace for XML Schema Instance
  xmlns_xsd: string; // Namespace for XML Schema Definition
  xmlns: string; // Default namespace for the WebPart
  ID: string; // ID attribute as a string
  Title?: string; // Title of the WebPart
  FrameType?: "" | "None" | "TitleAndBorder" | "TitleOnly" | "BorderOnly"; // Type of frame
  Description?: string; // Optional description
  IsIncluded: boolean; // Whether the WebPart is included
  ZoneID?: number; // Zone ID where the WebPart is placed
  PartOrder?: number; // Order of the part in the zone
  FrameState?: string; // State of the frame (e.g., Normal, Minimized)
  Height?: string; // Optional height
  Width?: string; // Optional width
  AllowRemove: boolean; // Whether the part can be removed
  AllowZoneChange: boolean; // Whether the part can change zones
  AllowMinimize: boolean; // Whether the part can be minimized
  AllowConnect: boolean; // Whether the part can connect to other parts
  AllowEdit: boolean; // Whether the part can be edited
  AllowHide: boolean; // Whether the part can be hidden
  IsVisible: boolean; // Whether the part is visible
  DetailLink?: string; // Optional detail link
  HelpLink?: string; // Optional help link
  HelpMode: string; // Help display mode (e.g., Modeless)
  Dir?: "" | "None" | "RightToLeft" | "LeftToRight"; // Direction (e.g., Default)
  PartImageSmall?: string; // Optional small part image
  MissingAssembly?: string; // Missing assembly message
  PartImageLarge?: string; // Optional large part image
  IsIncludedFilter?: string; // Optional filter
  ExportControlledProperties: boolean; // Export controlled properties flag
  ConnectionID: string; // Connection ID
  ID_: string; // ID with different format
  Assembly?: string; // Assembly reference
  TypeName?: string; // Type of the WebPart
  HeaderTitle?: string; // Title of the header (under a different namespace)
};

export default WebPartProperties;
