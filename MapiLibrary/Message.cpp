﻿#include "pch.h"

#include "Message.h"
#include "sha256.h"

namespace MapiLibrary
{
	SizedSPropTagArray(53, messageTags) =
	{
		53,
		{
			PR_ACCESS, PR_ACCESS_LEVEL, PR_BODY, PR_CLIENT_SUBMIT_TIME,
			PR_CONVERSATION_INDEX, PR_CREATION_TIME, PR_DISPLAY_NAME,
			PR_SUBJECT, PR_SENT_REPRESENTING_NAME, PR_MESSAGE_DELIVERY_TIME,
			PR_DISPLAY_BCC, PR_DISPLAY_CC, PR_DISPLAY_TO, PR_HASATTACH,
			PR_HTML, PR_IMPORTANCE, PR_INTERNET_CPID,
			PR_LAST_MODIFICATION_TIME, PR_MAPPING_SIGNATURE, PR_MDB_PROVIDER,
			PR_MESSAGE_ATTACHMENTS, PR_MESSAGE_CLASS, PR_MESSAGE_DELIVERY_TIME,
			PR_MESSAGE_FLAGS, PR_MESSAGE_RECIPIENTS, PR_NORMALIZED_SUBJECT,
			PR_OBJECT_TYPE, PR_RECORD_KEY, 	PR_RTF_COMPRESSED, PR_RTF_IN_SYNC,
			PR_RECEIVED_BY_ADDRTYPE, PR_RECEIVED_BY_EMAIL_ADDRESS,
			PR_RECEIVED_BY_ENTRYID, PR_RECEIVED_BY_NAME,
			PR_RECEIVED_BY_SEARCH_KEY, PR_REPLY_RECIPIENT_ENTRIES,
			PR_REPLY_RECIPIENT_NAMES, PR_SEARCH_KEY, PR_SENDER_ADDRTYPE,
			PR_SENDER_EMAIL_ADDRESS, PR_SENDER_NAME,
			PR_SENT_REPRESENTING_ADDRTYPE, PR_SENT_REPRESENTING_EMAIL_ADDRESS,
			PR_SENT_REPRESENTING_NAME, PR_SUBJECT_PREFIX, PR_SUBJECT,
			PR_INTERNET_MESSAGE_ID, PR_SENDER_ENTRYID, PR_SENDER_SEARCH_KEY,
			PR_SENT_REPRESENTING_ENTRYID, PR_SENT_REPRESENTING_NAME,
			PR_SENT_REPRESENTING_SEARCH_KEY, PR_TRANSPORT_MESSAGE_HEADERS,
		}
	};

	Message::Message(LPMESSAGE messageIn)
		: message(messageIn)
	{

	}

	std::string Message::GetMessageHash()
	{
		std::string base64Hash;

		unsigned long values;
		LPSPropValue messageProperties;
		HRESULT result = message->GetProps(
			(LPSPropTagArray)&messageTags,
			0,
			&values,
			&messageProperties);

		if (result == S_OK || result == MAPI_W_ERRORS_RETURNED)
		{
			SPropValue property = messageProperties[8];
			const std::wstring subjectWide(property.Value.lpszW);

			size_t size = subjectWide.size();

			int size_needed = WideCharToMultiByte(CP_UTF8, 0, &subjectWide[0],
				(int)size, NULL, 0, NULL, NULL);
			std::string subject(size_needed, 0);
			WideCharToMultiByte(CP_UTF8, 0, &subjectWide[0],
				(int)size, &subject[0], size_needed, NULL, NULL);

			std::cout << "subject: " << subject << std::endl;

			//std::vector<byte> hash;
			//size = wcsnlen(property.Value.lpszW, 256) * 2;

			//byte* bytes = (byte*)property.Value.lpszW;
			//hash.resize(size);

			//byte* end = bytes + size;

			////copy(bytes, end, back_inserter(hash));
			//hash.insert(hash.begin(), bytes, end);

			base64Hash = sha256(subject);
		}

		return base64Hash;
	}
}
