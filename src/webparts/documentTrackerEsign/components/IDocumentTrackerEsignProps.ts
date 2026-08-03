import { SPFI } from "@pnp/sp";

export interface IDocumentTrackerEsignProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: SPFI;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  context: any;
}
