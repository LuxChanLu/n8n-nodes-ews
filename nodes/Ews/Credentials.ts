import {
	ICredentialDataDecryptedObject,
} from 'n8n-workflow';
import { ExchangeService, ExchangeVersion, WebCredentials, Uri, ConfigurationApi } from 'ews-javascript-api'
// @ts-nocheck
const { ntlmAuthXhrApi } = require('ews-javascript-api-auth')

export function buildEwsWithNtlm(credentials: ICredentialDataDecryptedObject | undefined): ExchangeService {
	const exchange = new ExchangeService(credentials?.ewsVersion as ExchangeVersion)
	exchange.Credentials = new WebCredentials(credentials?.username as string,credentials?.password as string)
	exchange.Url = new Uri(`${credentials?.url}/EWS/Exchange.asmx`)
	ConfigurationApi.ConfigureXHR(new ntlmAuthXhrApi(credentials?.username as string,credentials?.password as string, true))
	return exchange
}
