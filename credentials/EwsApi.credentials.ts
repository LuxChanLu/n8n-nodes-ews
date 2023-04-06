import type { ICredentialType, INodeProperties } from 'n8n-workflow';
import { ExchangeVersion } from 'ews-javascript-api';

const ExchangeNames = Object.keys(ExchangeVersion).filter(k => typeof ExchangeVersion[k as any] === "number")

export class EwsApi implements ICredentialType {
	name = 'ewsApi';

	displayName = 'EWS - NTLM API';

	properties: INodeProperties[] = [
		{
			displayName: 'URL',
			name: 'url',
			required: true,
			type: 'string',
			default: '',
			placeholder: 'https://ews.localhost',
		},
		{
			displayName: 'Exchange version',
			name: 'ewsVersion',
			type: 'options',
			default: ExchangeNames[0],
			options: ExchangeNames.map(name => ({ name, value: ExchangeVersion[name as any] }))
		},
		{
			displayName: 'Username',
			name: 'username',
			type: 'string',
			default: '',
		},
		{
			displayName: 'Password',
			name: 'password',
			type: 'string',
			typeOptions: {
				password: true,
			},
			default: '',
		},
	];
}
