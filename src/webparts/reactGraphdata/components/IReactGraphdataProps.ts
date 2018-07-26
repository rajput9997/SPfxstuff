
import { IReactGraphdataWebPartProps } from "../ReactGraphdataWebPart";
import { DisplayMode } from "@microsoft/sp-core-library";
import { MSGraphClient } from "@microsoft/sp-client-preview";

export interface IReactGraphdataProps extends IReactGraphdataWebPartProps {
  displayMode: DisplayMode;
  graphClient: MSGraphClient;
  updateProperty: (value: string) => void;
}