import type { INodeProperties } from 'n8n-workflow';

export const attachmentOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		default: 'get',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['attachment'],
			},
		},
		options: [
			{
				name: 'Get',
				value: 'get',
				description: 'Get an attachment',
				action: 'Get an attachment',
			},
		],
	},
];


export const attachmentDescription: INodeProperties[] = [
	/* -------------------------------------------------------------------------- */
	/*                                attachment:get                              */
	/* -------------------------------------------------------------------------- */
	{
		displayName: 'Attachment ID',
		name: 'attachmentId',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				operation: ['get'],
				resource: ['attachment'],
			},
		},
	},
];
