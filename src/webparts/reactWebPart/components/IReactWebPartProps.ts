import { MSGraphClient } from "@microsoft/sp-http";

export interface IReactWebPartProps {
  description: string;
  graphClient:MSGraphClient;
}
