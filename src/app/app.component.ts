import { NONE_TYPE } from '@angular/compiler';
import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  id = "";

  func(){
    var mail = Office.context.mailbox.item;
    this.id = mail.itemId;
    console.log(mail);
    // var request = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'+
    //           '  <soap:Header><t:RequestServerVersion Version="Exchange2010" /></soap:Header>'+
    //           '  <soap:Body>'+
    //           '    <m:CreateItem MessageDisposition="SendAndSaveCopy">'+
    //           '      <m:SavedItemFolderId><t:DistinguishedFolderId Id="sentitems" /></m:SavedItemFolderId>'+
    //           '      <m:Items>'+
    //           '        <t:Message>'+
    //           '          <t:Subject>Hello, Outlook!</t:Subject>'+
    //           '          <t:Body BodyType="HTML">Hello World!</t:Body>'+
    //           '          <t:ToRecipients>'+
    //           '            <t:Mailbox><t:EmailAddress>' + Office.context.mailbox.userProfile.emailAddress + '</t:EmailAddress></t:Mailbox>'+
    //           '          </t:ToRecipients>'+
    //           '        </t:Message>'+
    //           '      </m:Items>'+
    //           '    </m:CreateItem>'+
    //           '  </soap:Body>'+
    //           '</soap:Envelope>';

    //         Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
    //           if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    //             console.log("Action failed with error: " + asyncResult.error.message);
    //           }
    //           else {
    //             console.log("Message sent!");
    //           }
    //         });
  }

}
