import { MSGraphClient } from '@microsoft/sp-http';

export interface IMsgPeronaProps {
  graphClient: MSGraphClient;
}



export interface IGraphPersonaProps {
  graphClient: MSGraphClient;
}

export interface IGraphPersonaState {
  name: string;
  email: string;
  phone: string;
  image: string;
}