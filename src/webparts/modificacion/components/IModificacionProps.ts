import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IModificacionProps {
  context: WebPartContext;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
