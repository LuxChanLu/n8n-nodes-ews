import { BasePropertySet, ConflictResolutionMode, EmailMessageSchema, ExchangeService, FindItemsResults, Item, ItemId, ItemView, PropertySet, WellKnownFolderName } from "ews-javascript-api"
import { IDataObject, IExecuteFunctions, INodeExecutionData } from "n8n-workflow"
import { safeGetter } from "../utils"

export async function find(func: IExecuteFunctions, exchange: ExchangeService, idx: number, returnData: INodeExecutionData[]): Promise<any> {
	const folder = func.getNodeParameter('wellKnownFolderName', idx) as WellKnownFolderName
	const searchQuery = func.getNodeParameter('searchQuery', idx) as string
	const itemCount = func.getNodeParameter('itemCount', idx) as number

	returnData.push({ json: transformFindResult(await exchange.FindItems(folder, searchQuery, new ItemView(itemCount))), pairedItem: idx })
}

export async function get(func: IExecuteFunctions, exchange: ExchangeService, idx: number, returnData: INodeExecutionData[]): Promise<any> {
	const itemId = func.getNodeParameter('itemId', idx) as string
	const items = await exchange.BindToItems([new ItemId(itemId)], new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments))
	for (const response of items.Responses) {
		returnData.push({ json: transformItem(response.Item), pairedItem: idx })
	}
}

export async function setCategory(func: IExecuteFunctions, exchange: ExchangeService, idx: number, returnData: INodeExecutionData[]): Promise<any> {
	const itemId = func.getNodeParameter('itemId', idx) as string
	const category = func.getNodeParameter('category', idx) as string
	const items = await exchange.BindToItems([new ItemId(itemId)], new PropertySet(BasePropertySet.FirstClassProperties))
	const item = items.Responses.at(0)?.Item
	if (item && !item.Categories.Contains(category)) {
		item.Categories.Add(category)
		await item.Update(ConflictResolutionMode.AutoResolve)
		returnData.push({ json: transformItem(item), pairedItem: idx })
	}
}

const transformItem = (item: Item): IDataObject => {
	return {
		id: safeGetter(() => item.Id.UniqueId),
		isNew: safeGetter(() => item.IsNew),
		mimeContent: safeGetter(() => item.MimeContent),
		parentFolderId: safeGetter(() => item.ParentFolderId.UniqueId),
		sensitivity: safeGetter(() => item.Sensitivity),
		dateTimeReceived: safeGetter(() => item.DateTimeReceived.ToISOString()),
		size: safeGetter(() => item.Size),
		categories: safeGetter(() => item.Categories.GetEnumerator()),
		culture: safeGetter(() => item.Culture),
		importance: safeGetter(() => item.Importance),
		inReplyTo: safeGetter(() => item.InReplyTo),
		isSubmitted: safeGetter(() => item.IsSubmitted),
		isAssociated: safeGetter(() => item.IsAssociated),
		isDraft: safeGetter(() => item.IsDraft),
		isFromMe: safeGetter(() => item.IsFromMe),
		isResend: safeGetter(() => item.IsResend),
		isUnmodified: safeGetter(() => item.IsUnmodified),
		internetMessageHeaders: safeGetter(() => item.InternetMessageHeaders.GetEnumerator().map(header => ({ name: header.Name, value: header.Value }))),
		dateTimeSent: safeGetter(() => item.DateTimeSent?.ToISOString()),
		dateTimeCreated: safeGetter(() => item.DateTimeCreated?.ToISOString()),
		allowedResponseActions: safeGetter(() => item.AllowedResponseActions),
		reminderDueBy: safeGetter(() => item.ReminderDueBy),
		isReminderSet: safeGetter(() => item.IsReminderSet),
		reminderMinutesBeforeStart: safeGetter(() => item.ReminderMinutesBeforeStart),
		displayCc: safeGetter(() => item.DisplayCc),
		displayTo: safeGetter(() => item.DisplayTo),
		hasAttachments: safeGetter(() => item.HasAttachments),
		attachments: safeGetter(() => item.Attachments.GetEnumerator().map(attachment => ({
			id: attachment.Id,
			name: attachment.Name,
			contentType: attachment.ContentType,
			contentId: attachment.ContentId,
			contentLocation: attachment.ContentLocation,
			size: attachment.Size,
			lastModifiedTime: attachment.LastModifiedTime.ToISOString(),
			isInline: attachment.IsInline,
		}))),
		body: safeGetter(() => ({ bodyType: item.Body.BodyType, text: item.Body.Text })),
		itemClass: safeGetter(() => item.ItemClass),
		subject: safeGetter(() => item.Subject),
		webClientReadFormQueryString: safeGetter(() => item.WebClientReadFormQueryString),
		webClientEditFormQueryString: safeGetter(() => item.WebClientEditFormQueryString),
		extendedProperties: safeGetter(() => item.ExtendedProperties.GetEnumerator()),
		effectiveRights: safeGetter(() => item.EffectiveRights),
		lastModifiedName: safeGetter(() => item.LastModifiedName),
		lastModifiedTime: safeGetter(() => item.LastModifiedTime?.ToISOString()),
		conversationId: safeGetter(() => item.ConversationId.UniqueId),
		uniqueBody: safeGetter(() => item.UniqueBody),
		storeEntryId: safeGetter(() => item.StoreEntryId),
		instanceKey: safeGetter(() => item.InstanceKey),
		flag: safeGetter(() => ({
			status: item.Flag.FlagStatus,
			startDate: item.Flag.StartDate?.ToISOString(),
			dueDate: item.Flag.DueDate?.ToISOString(),
			completeDate: item.Flag.CompleteDate?.ToISOString()
		})),
		normalizedBody: safeGetter(() => item.NormalizedBody),
		entityExtractionResult: safeGetter(() => ({
			addresses: item.EntityExtractionResult.Addresses.GetEnumerator().map(address => ({ address: address.Address })),
			meetingSuggestions: item.EntityExtractionResult.MeetingSuggestions.GetEnumerator().map(meetingSuggestion => ({
				attendees: meetingSuggestion.Attendees.GetEnumerator().map(attendee => ({ name: attendee.Name, userId: attendee.UserId })),
				location: meetingSuggestion.Location,
				subject: meetingSuggestion.Subject,
				meetingString: meetingSuggestion.MeetingString,
				startTime: meetingSuggestion.StartTime.ToISOString(),
				endTime: meetingSuggestion.EndTime.ToISOString(),
			})),
			taskSuggestions: item.EntityExtractionResult.TaskSuggestions.GetEnumerator().map(taskSuggestion => ({
				taskString: taskSuggestion.TaskString,
				assignees: taskSuggestion.Assignees?.GetEnumerator().map(assignee => ({ name: assignee.Name, userId: assignee.UserId }))
			})),
			emailAddresses: item.EntityExtractionResult.EmailAddresses.GetEnumerator().map(emailAddress => ({ emailAddress: emailAddress.EmailAddress })),
			contacts: item.EntityExtractionResult.Contacts.GetEnumerator().map(contact => ({
				personName: contact.PersonName,
				businessName: contact.BusinessName,
				phoneNumbers: contact.PhoneNumbers?.GetEnumerator().map(phoneNumber => ({
					originalPhoneString: phoneNumber.OriginalPhoneString,
					phoneString: phoneNumber.PhoneString,
					type: phoneNumber.Type
				})),
				urls: contact.Urls?.GetEnumerator(),
				emailAddresses: contact.EmailAddresses?.GetEnumerator(),
				addresses: contact.Addresses?.GetEnumerator(),
				contactString: contact.ContactString
			})),
			urls: item.EntityExtractionResult.Urls.GetEnumerator().map(url => ({ position: url.Position, url: url.Url })),
			phoneNumbers: item.EntityExtractionResult.PhoneNumbers.GetEnumerator().map(phoneNumber => ({
				Position: phoneNumber.Position,
				OriginalPhoneString: phoneNumber.OriginalPhoneString,
				PhoneString: phoneNumber.PhoneString,
				Type: phoneNumber.Type
			}))
		})),
		policyTag: safeGetter(() => item.PolicyTag),
		archiveTag: safeGetter(() => item.ArchiveTag),
		retentionDate: safeGetter(() => item.RetentionDate?.ToISOString()),
		preview: safeGetter(() => item.Preview),
		textBody: safeGetter(() => item.TextBody),
		iconIndex: safeGetter(() => item.IconIndex),
	}
}

const transformFindResult = (messages: FindItemsResults<Item>): IDataObject => {
	return {
		moreAvailable: messages.MoreAvailable,
		nextPageOffset: messages.NextPageOffset,
		totalCount: messages.TotalCount,
		highlightTerms: messages.HighlightTerms.map(term => ({ scope: term.Scope, value: term.Value })),
		items: messages.Items.map(transformItem)
	}
}
