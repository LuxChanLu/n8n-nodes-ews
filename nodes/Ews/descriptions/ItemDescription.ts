import type { INodeProperties } from 'n8n-workflow';
import { WellKnownFolderName } from 'ews-javascript-api';

const WellKnownFolderNames = Object.keys(WellKnownFolderName).filter(k => typeof WellKnownFolderName[k as any] === "number")

export const itemOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		default: 'find',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['item'],
			},
		},
		options: [
			{
				name: 'Find',
				value: 'find',
				description: 'Find an item',
				action: 'Find an item',
			},
		],
	},
];


export const itemDescription: INodeProperties[] = [
	/* -------------------------------------------------------------------------- */
	/*                                   item:find                                */
	/* -------------------------------------------------------------------------- */
	{
		displayName: 'Search Query (Advanced Query Syntax (AQS))',
		name: 'searchQuery',
		type: 'string',
		default: '',
		placeholder: 'from:',
		required: true,
		displayOptions: {
			show: {
				operation: ['find'],
				resource: ['item'],
			},
		},
	},
	{
		displayName: 'Well Known Folder',
		name: 'wellKnownFolderName',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				operation: ['find'],
				resource: ['item'],
			},
		},
		default: WellKnownFolderNames[0],
		options: WellKnownFolderNames.map(name => ({ name, value: WellKnownFolderName[name as any] }))
	},
	{
		displayName: 'Item count',
		name: 'itemCount',
		type: 'number',
		default: 100,
		required: true,
		displayOptions: {
			show: {
				operation: ['find'],
				resource: ['item'],
			},
		},
	},
];
