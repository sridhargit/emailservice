package com.transunion.wf;

import com.transunion.wf.exception.MailActionsException;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import org.apache.commons.lang3.StringUtils;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Map;

public class MailActions {

	private ExchangeService exchangeService;
	private String mailEWSURL;
	private ExchangeCredentials credentials;

	public MailActions(String mailEWSURL, ExchangeCredentials credentials) throws URISyntaxException {
		this.mailEWSURL = mailEWSURL;
		this.credentials = credentials;
		init();
	}

	private void init() throws URISyntaxException {
		exchangeService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
		exchangeService.setUrl(new URI(mailEWSURL));
		exchangeService.setCredentials(credentials);
	}

	/**
	 * This method accepts the multiple attachments in map with attachment name and
	 * attachment as byte array and sends an email.
	 *
	 * @param recipients
	 * @param ccRecipients
	 * @param subject
	 * @param messageBody
	 * @param attachments
	 */
	public void sendEmailWithMultipleAttachments(String recipients, String ccRecipients, String subject, MessageBody messageBody,
												 Map<String,byte[]> attachments){
		try {
			EmailMessage message = new EmailMessage(exchangeService);
			message.setSubject(subject);
			message.setBody(messageBody);
			getToRecipients(recipients, message);
			if(!StringUtils.isEmpty(ccRecipients)) {
				getCCRecipients(ccRecipients, message);
			}
			AttachmentCollection attachmentCollection = message.getAttachments();
			for(Map.Entry<String,byte[]> entry : attachments.entrySet()){
				attachmentCollection.addFileAttachment(entry.getKey(), entry.getValue());
			}
			message.send();
		} catch (Exception ex) {
			throw new MailActionsException(ex,"Unable to Send a Mail with multi attachments" + ex.getMessage());
		}

	}

	public void sendEmails(String recipients, String ccRecipients, String subject, MessageBody messageBody,
			byte[] attachmentBytes, String attachmentName) throws Exception {
		try {
			EmailMessage message = new EmailMessage(exchangeService);
			message.setSubject(subject);
			message.setBody(messageBody);
			getToRecipients(recipients, message);
			if(!StringUtils.isEmpty(ccRecipients)) {
				getCCRecipients(ccRecipients, message);
				}
			message.getAttachments().addFileAttachment(attachmentName, attachmentBytes);
			message.send();
		} catch (Exception ex) {
			throw new MailActionsException("Unable to Send a Mail with attachment" + ex.getMessage());
		}
	}

	public void sendEmailsWithoutAttachment(String recipients, String ccRecipients, String subject,
			MessageBody messageBody) throws Exception {
		try {
			EmailMessage message = new EmailMessage(exchangeService);
			message.setSubject(subject);
			message.setBody(messageBody);
			getToRecipients(recipients, message);
			getCCRecipients(ccRecipients, message);
			message.send();
		} catch (Exception ex) {
			throw new MailActionsException("Unable to Send a Mail with attachment" + ex.getMessage());
		}
	}


	private EmailMessage getToRecipients(String recipients, EmailMessage message) throws ServiceLocalException {
		if(recipients!=null && !"".equals(recipients)) {
			String[] addresses = recipients.split(",");
			for (String address : addresses) {
				message.getToRecipients().add(address);
			}
		}
		return message;
	}

	private EmailMessage getCCRecipients(String ccrecipients, EmailMessage message) throws ServiceLocalException {
		if(ccrecipients!=null && !"".equals(ccrecipients)) {
			 String[] addresses = ccrecipients.split(",");
			for (String address : addresses) {
				message.getCcRecipients().add(address);
			}
		}
		return message;
	}
}
