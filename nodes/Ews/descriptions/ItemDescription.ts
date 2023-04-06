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
			{
				name: 'Get',
				value: 'get',
				description: 'Get an item',
				action: 'Get an item',
			},
			{
				name: 'Set categories',
				value: 'set-categories',
				description: 'Set an item categories',
				action: 'Set an item categories',
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
	/* -------------------------------------------------------------------------- */
	/*                                  item:get                                  */
	/* -------------------------------------------------------------------------- */
	{
		displayName: 'Item ID',
		name: 'itemId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				operation: ['get'],
				resource: ['item'],
			},
		},
	},
	/* -------------------------------------------------------------------------- */
	/*                             item:set-categories                            */
	/* -------------------------------------------------------------------------- */
	{
		displayName: 'Item ID',
		name: 'itemId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				operation: ['set-categories'],
				resource: ['item'],
			},
		},
	},
	{
		displayName: 'Categories',
		name: 'categories',
		type: 'collection',
		default: '',
		required: true,
		displayOptions: {
			show: {
				operation: ['set-categories'],
				resource: ['item'],
			},
		},
	},
];
