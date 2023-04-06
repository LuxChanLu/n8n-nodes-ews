import {
	ICredentialTestFunctions,
	INodeCredentialTestResult,
	ICredentialsDecrypted,
	ICredentialDataDecryptedObject,
} from 'n8n-workflow';
import { ExchangeService, ExchangeVersion, WebCredentials, Uri, ConfigurationApi } from 'ews-javascript-api'
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-nocheck
const { ntlmAuthXhrApi } = require('ews-javascript-api-auth')

export function buildEwsWithNtlm(credentials: ICredentialDataDecryptedObject | undefined): ExchangeService {
	const exchange = new ExchangeService(credentials?.ewsVersion as ExchangeVersion)
	exchange.Credentials = new WebCredentials(credentials?.username as string,credentials?.password as string)
	exchange.Url = new Uri(`${credentials?.url}/EWS/Exchange.asmx`)
	ConfigurationApi.ConfigureXHR(new ntlmAuthXhrApi(credentials?.username as string,credentials?.password as string, true))
	return exchange
}

export async function ewsNtlmApiTest(this: ICredentialTestFunctions, credential: ICredentialsDecrypted): Promise<INodeCredentialTestResult> {
	try {
		const exchange = buildEwsWithNtlm(credential.data)

		await exchange.GetServerTimeZones()
	} catch (error) {
		return {
			status: 'Error',
			message: `Settings are not valid or authentification failed: ${error}`,
		};
	}
	return {
		status: 'OK',
		message: 'Authentication successful!',
	};
}
