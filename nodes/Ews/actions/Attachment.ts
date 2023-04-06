import { BodyType, ExchangeService, FileAttachment } from "ews-javascript-api"
import { IExecuteFunctions, INodeExecutionData } from "n8n-workflow"

export async function get(func: IExecuteFunctions, exchange: ExchangeService, idx: number, returnData: INodeExecutionData[]): Promise<any> {
	const attachmentId = func.getNodeParameter('attachmentId', idx) as string
	const items = await exchange.GetAttachments([attachmentId], BodyType.Text, [])
	const attachment = items.Responses.at(0)?.Attachment as FileAttachment
	if (attachment) {
		returnData.push({
			json: {
				id: attachment.Id,
				name: attachment.Name,
				contentType: attachment.ContentType,
				contentId: attachment.ContentId,
				contentLocation: attachment.ContentLocation,
				size: attachment.Size,
				lastModifiedTime: attachment.LastModifiedTime,
				isInline: attachment.IsInline,
				fileName: attachment.FileName,
				isContactPhoto: attachment.IsContactPhoto,
			},
			binary: { attachment: await func.helpers.prepareBinaryData(Buffer.from(attachment.Base64Content, 'base64'), attachment.FileName), },
			pairedItem: idx
		})
	}
}
