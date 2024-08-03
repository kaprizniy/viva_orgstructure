import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IVivaOrgstructureWebPartProps {
	description: string;
	employeeListLink: string;
	startLogin: string;
	loginDomain: string;
	context: WebPartContext;
}