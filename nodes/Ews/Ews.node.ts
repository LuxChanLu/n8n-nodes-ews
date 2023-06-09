import {
	IExecuteFunctions,
} from 'n8n-core';

import {
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	deepCopy,
	NodeOperationError,
	ICredentialTestFunctions,
	ICredentialsDecrypted,
	INodeCredentialTestResult,
} from 'n8n-workflow';

import {
	itemDescription, itemOperations,
	attachmentDescription, attachmentOperations
} from './descriptions';

import {
	buildEwsWithNtlm,
} from './Credentials';

import * as Actions from './actions';
import { ExchangeService } from 'ews-javascript-api'


export class Ews implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Exchange Web Service',
		name: 'ews',
		icon: 'file:ews.svg',
		group: ['output'],
		version: 1,
		description: 'Exchange Web Service API',
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		defaults: {
			name: 'Exchange Web Service',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'ewsApi',
				required: true,
				testedBy: 'ewsNtlmApiTest',
				displayOptions: {
					show: {
						authentication: ['ntlm'],
					},
				},
			},
		],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				default: 'item',
				noDataExpression: true,
				options: [
					{
						name: 'Item',
						value: 'item',
					},
					{
						name: 'Attachment',
						value: 'attachment',
					},
				],
			},
			{
				displayName: 'Authentication',
				name: 'authentication',
				type: 'options',
				options: [
					{
						name: 'NTLM',
						value: 'ntlm',
					},
				],
				default: 'ntlm',
				description: 'Means of authenticating with the service',
			},


			...itemOperations, ...itemDescription,
			...attachmentOperations, ...attachmentDescription
		]
	};

	methods = {
		credentialTest: {
			async ewsNtlmApiTest(this: ICredentialTestFunctions, credential: ICredentialsDecrypted): Promise<INodeCredentialTestResult> {
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
		}
	};

	// The execute method will go here
	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		let items = this.getInputData();
		items = deepCopy(items);

		const resource = this.getNodeParameter('resource', 0);
		const operation = this.getNodeParameter('operation', 0);

		const returnData: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; ++i) {
			try {
				await (Actions as any)[resource][operation](this, await exchangeService(this, i), i, returnData)
			} catch (error) {
				if (this.continueOnFail()) {
					const executionData = this.helpers.constructExecutionMetaData(
						this.helpers.returnJsonArray({ error: error.message }),
						{ itemData: { item: i } },
					);
					returnData.push(...executionData);

					continue;
				}
				throw error;
			}
		}
		return this.prepareOutputData(returnData);
	};
}

async function exchangeService(func: IExecuteFunctions, idx: number): Promise<ExchangeService> {
	const authenticationMethod = func.getNodeParameter('authentication', idx) as string;
	switch (authenticationMethod) {
		case 'ntlm':
			return buildEwsWithNtlm(await func.getCredentials('ewsApi', idx));
		default:
			break;
	}
	throw new NodeOperationError(func.getNode(), new Error('No authentification method found'), { itemIndex: idx })
}
