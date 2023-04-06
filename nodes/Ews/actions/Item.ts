import { ExchangeService, FindItemsResults, Item, ItemView, WellKnownFolderName } from "ews-javascript-api"
import { IDataObject, IExecuteFunctions, INodeExecutionData } from "n8n-workflow"

export async function find(func: IExecuteFunctions, exchange: ExchangeService, idx: number, returnData: INodeExecutionData[]): Promise<any> {
	const folder = func.getNodeParameter('wellKnownFolderName', idx) as WellKnownFolderName
	const searchQuery = func.getNodeParameter('searchQuery', idx) as string
	const itemCount = func.getNodeParameter('itemCount', idx) as number

	const transformResult = (messages: FindItemsResults<Item>): IDataObject => {
		return {
			moreAvailable: messages.MoreAvailable,
			nextPageOffset: messages.NextPageOffset,
			totalCount: messages.TotalCount,
			highlightTerms: messages.HighlightTerms.map(term => ({ scope: term.Scope, value: term.Value })),
			items: messages.Items.map(item => ({ ...item }))
		}
	}

	returnData.push({ json: transformResult(await exchange.FindItems(folder, searchQuery, new ItemView(itemCount))), pairedItem: idx })
}
