import { BodyType, ExchangeService, FileAttachment } from "ews-javascript-api"
import { IExecuteFunctions, INodeExecutionData } from "n8n-workflow"

export async function get(func: IExecuteFunctions, exchange: ExchangeService, idx: number, returnData: INodeExecutionData[]): Promise<any> {
	const attachmentId = func.getNodeParameter('attachmentId', idx) as string
	const items = await exchange.GetAttachments([attachmentId], BodyType.Text, [])
	const attachement = items.Responses.at(0)?.Attachment as FileAttachment
	if (attachement) {
		returnData.push({
			json: {
				id: attachement.Id,
				name: attachement.Name,
				contentType: attachement.ContentType,
				contentId: attachement.ContentId,
				contentLocation: attachement.ContentLocation,
				size: attachement.Size,
				lastModifiedTime: attachement.LastModifiedTime,
				isInline: attachement.IsInline,
				fileName: attachement.FileName,
				isContactPhoto: attachement.IsContactPhoto,
			},
			binary: { attachement: await func.helpers.prepareBinaryData(Buffer.from(attachement.Base64Content, 'base64'), attachement.FileName), },
			pairedItem: idx
		})
	}
}
