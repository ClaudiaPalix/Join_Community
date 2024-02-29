//import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IYammerProvider } from "../../../utils/yammer/IYammerProvider";

export interface IJoinCommunityProps {
  context: any;
  yammerProvider: IYammerProvider;
  selectedList: string;
  selectedList2: string;
  userEmail: any;
  seeAllUrl: string;
  description: string;
  isTeams: boolean;
  isEmbedded: boolean;
}
