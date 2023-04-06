import { BodyType, ExchangeService } from "ews-javascript-api"
import { IExecuteFunctions, INodeExecutionData } from "n8n-workflow"

export async function get(func: IExecuteFunctions, exchange: ExchangeService, idx: number, returnData: INodeExecutionData[]): Promise<any> {
	const attachementId = func.getNodeParameter('attachementId', idx) as string
	const items = await exchange.GetAttachments([attachementId], BodyType.Text, [])
	console.log(items.Responses)
}
