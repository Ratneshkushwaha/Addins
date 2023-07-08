/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import {
    allComponents,
    provideFluentDesignSystem,
  } from "@fluentui/web-components";
  provideFluentDesignSystem().register(allComponents);
  
  let AccessToken = "";
  let RefreshToken = "";
  let outlookToken = "";
  let ApiUrl = "https://utility-1.grunley.info/OutlookWebAPI/api/";
  let FilterData = [];
  let EmailResult;
  let TagTextValue;
  $(document).ready(() => {});
  
  // The initialize function must be run each time a new page is loaded
  Office.initialize = async (reason) => {
    $("#sideload-msg").hide();
    $("#app-body").show();
    $("#Email").append(Office.context.mailbox.userProfile.emailAddress);
    getDrpDepartment();
    getAccessToken();
    requestToken();
    //getSelectedMailAttchments();
    // Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    //   if (result.status !== Office.AsyncResultStatus.Succeeded) {
    //     console.error(
    //       `Token retrieval failed with message: ${result.error.message}`
    //     );
    //   } else {
    //     // Use the outlook access token.
    //     outlookToken = result.value;
  
    //   }
    // });
    let url = "https://graph.microsoft.com/v1.0/me/outlook/masterCategories";
    $.ajax({
      url: url,
      type: "POST",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      data: JSON.stringify({
        displayName: "Project",
        color: "preset9",
      }),
      success: function (categoryResult) {
        console.log(categoryResult);
      },
      error: function (error) {
        console.log("Error in getting data: category Result " + error);
      },
    });
  };
  
  function requestToken() {
    $.ajax({
      async: false,
      crossDomain: true,
      // "url": "https://login.microsoftonline.com/0a6bce9c-73e9-4d31-88fc-240e3069abfe/oauth2/token",
      url: "https://howling-crypt-47129.herokuapp.com/https://login.microsoftonline.com/0a6bce9c-73e9-4d31-88fc-240e3069abfe/oauth2/v2.0/token", // Pass your tenant name instead of sharepointtechie
      method: "POST",
      headers: {
        "content-type": "application/x-www-form-urlencoded",
      },
      data: {
        grant_type: "refresh_token",
        "client_id ": "7b8ee4b7-4c4f-4e6b-88a7-b8f30fe849e2", //Provide your app id
        client_secret: "0Muvy2qb6[.Mh?fRQ[ErKQCrwtpF0GUV", //Provide your secret
        refresh_token: RefreshToken,
        "scope ": "https://graph.microsoft.com/.default",
      },
      success: function (response) {
        // console.log(response);
        AccessToken = response.access_token; //Store the token into global variable
      },
      error: function (error) {
        console.log(JSON.stringify(error));
      },
    });
  }
  function getAccessToken() {
    //get access Token
    const requestUrl = ApiUrl + "Addin/GetAccessToken";
    $.ajax({
      url: requestUrl,
      async: false,
      dataType: "json",
    })
      .done(function (item) {
        AccessToken = item.Value.AccessToken;
        RefreshToken = item.Value.RefreshToken;
      })
      .fail(function (error) {
        console.log(error);
      });
  }
  function getDrpDepartment() {
    const Email = getEmail();
    const getMessageUrl1 = ApiUrl + "addin/GetDepartment?Email=" + Email;
    $.ajax({
      url: getMessageUrl1,
      dataType: "json",
      async: false,
    })
      .done(function (item) {
        let data = item.Value;
        data.forEach((element) => {
          $("#Department").append(
            `<fluent-option value='${element.Title}'>${element.Title}</fluent-option>`
          );
        });
        //console.log(item);
      })
      .fail(function (error) {
        console.log(error);
      });
  }
  function getCurrentItem(itemId) {
    let response;
    // const getMessageUrl =
    //   Office.context.mailbox.restUrl + "/v2.0/me/messages/" + itemId;
    const getMessageUrl =
      "https://graph.microsoft.com/v1.0/me/messages/" + itemId;
    $.ajax({
      type: "GET",
      contentType: "application/json",
      async: false,
      url: getMessageUrl,
      dataType: "json",
      headers: { Authorization: "Bearer " + AccessToken },
    })
      .done(function (item) {
        console.log(item);
        response = item;
      })
      .fail(function (error) {
        // Handle error.
        console.log(error);
        return [];
      });
    return response;
  }
  function getSelectedMailAttachmentItem(itemId) {
    let response;
    // const getMessageUrl =
    //   Office.context.mailbox.restUrl + "/v2.0/me/messages/" + itemId;
    const getMessageUrl =
      "https://graph.microsoft.com/v1.0/me/messages/" + itemId + "/attachments";
    $.ajax({
      type: "GET",
      contentType: "application/json",
      async: false,
      url: getMessageUrl,
      dataType: "json",
      headers: { Authorization: "Bearer " + AccessToken },
    })
      .done(function (item) {
        console.log(item);
        response = item;
      })
      .fail(function (error) {
        // Handle error.
        console.log(error);
        return [];
      });
    return response;
  }
  function getItemRestId() {
    if (Office.context.mailbox.diagnostics.hostName === "OutlookIOS") {
      // itemId is already REST-formatted.
      return Office.context.mailbox.item.itemId;
    } else {
      // Convert to an item ID for API v2.0.
      return Office.context.mailbox.convertToRestId(
        Office.context.mailbox.item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
    }
  }
  function getMIMEData(id) {
    let response = "";
    //const getMessageUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + id + "/$value";
    const getMessageUrl =
      "https://graph.microsoft.com/v1.0/me/messages/" + id + "/$value";
    $.ajax({
      url: getMessageUrl,
      async: false,
      contentType: "text/html; charset=UTF-8",
      headers: { Authorization: "Bearer " + AccessToken },
    })
      .done(function (item) {
        response = item;
      })
      .fail(function (error) {
        console.log(error);
      });
    return response;
  }
  $(document).on("change", "#Department", function (event) {
    console.log(event.target.value);
    $(".progress-section").removeClass("show");
    $("#progessbar").val("0");
    $("#Project").empty();
    const Email = getEmail();
    console.log($("#Department").val());
    const getMessageUrl =
      ApiUrl +
      "addin/GetTags?Department=" +
      event.target.value +
      "&Email=" +
      Email;
    $.ajax({
      url: getMessageUrl,
      dataType: "json",
    })
      .done(function (item) {
        let data = item.Value;
        data.length > 0
          ? data.forEach((element) => {
              $("#Project").append(
                `<fluent-option data-title='${element.Title}' value='${element.ID}'>${element.Title}</fluent-option>`
              );
            })
          : "";
      })
      .fail(function (error) {});
  });
  
  $(document).on("change", "#Project", async function (event) {
    console.log(event.currentTarget._selectedOptions[0].innerText);
    TagTextValue = event.currentTarget._selectedOptions[0].innerText;
  });
  $(document).on("click", "#btnArchive", async function (event) {
    $(".progress-section").addClass("progress-section show");
    debugger;
    Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return [];
      }
      let selectedEmails = asyncResult.value;
      for (let index = 0; index < selectedEmails.length; index++) {
        const element = selectedEmails[index];
        let itemId = Office.context.mailbox.convertToRestId(
          element.itemId,
          Office.MailboxEnums.RestVersion.v2_0
        );
        //getSelectedMailAttchments(element);
        //console.log(element);
        let MIMIData = getMIMEData(itemId);
        let item = getCurrentItem(itemId);
        uploadAttachmentFile(item, itemId, MIMIData);
      }
    });
  });
  function getTimeStamp(ReceivedDateTimeDate) {
    var now = new Date(ReceivedDateTimeDate);
    return (
      now.getMonth() +
      1 +
      "" +
      now.getDate() +
      "" +
      now.getFullYear() +
      "" +
      now.getHours() +
      "" +
      (now.getMinutes() < 10 ? "0" + now.getMinutes() : now.getMinutes()) +
      "" +
      (now.getSeconds() < 10 ? "0" + now.getSeconds() : now.getSeconds())
    );
  }
  function getSelectedMailAttchments(mail) {
    debugger;
    Office.context.mailbox.item.getCallbackTokenAsync(
      { isRest: true },
      function (result) {
        if (result.status === "succeeded") {
          var accessToken = result.value;
          var itemIds = Office.context.mailbox.item.getSelectedDataAsync(
            Office.CoercionType.ItemIds,
            function (result) {
              if (result.status === "succeeded") {
                var itemIds = result.value;
                itemIds.forEach(function (itemId) {
                  var attachmentUrl =
                    Office.context.mailbox.restUrl +
                    "/v2.0/me/messages/" +
                    itemId +
                    "/attachments";
  
                  $.ajax({
                    url: attachmentUrl,
                    type: "GET",
                    headers: {
                      Authorization: "Bearer " + accessToken,
                    },
                    success: function (response) {
                      var attachments = response.value;
                      // Process attachments as needed
                      console.log(attachments);
                    },
                    error: function (error) {
                      console.error("Error retrieving attachments:", error);
                    },
                  });
                });
              } else {
                console.error("Error retrieving item IDs:", result.error);
              }
            }
          );
        } else {
          console.error("Error retrieving access token:", result.error);
        }
      }
    );
  }
  function uploadAttachmentFile(mailItem, MessageId, MIMEData) {
    $("#progessbar").val("20");
    // moveItem(mailItem);
    let CurentSelectedEmail = mailItem;
    let ProjectNumber = $("#Project").val();
    console.log(CurentSelectedEmail);
    if (!!CurentSelectedEmail) {
      $("#progessbar").val("30");
      let Department = $("#Department").val();
      let TagId = $("#Project").val();
      let FromEmail = CurentSelectedEmail.sender.emailAddress.address;
      let checkTagAssigne = CheckIfTagAssigned(
        Department,
        TagId,
        MessageId,
        FromEmail
      );
      if (!!mailItem && !checkTagAssigne) {
        let datestemp = getTimeStamp(CurentSelectedEmail.receivedDateTime);
        let subject = mailItem.subject != null ? mailItem.subject : "";
        debugger;
        if (subject.length > 25) {
          subject = subject.substring(0, 25);
        }
        let project = $("#Project fluent-option:selected").text();
  
        let MyEmail = getEmail();
        let strFileName =
          datestemp + subject.replace(/[^a-zA-Z0-9_.]/g, "") + ".msg";
        let SelectedDepartmentText = $("#Department").val();
        let date = new Date(CurentSelectedEmail.receivedDateTime);
        let monthNumber =
          date.getMonth().toString().length == 1
            ? "0" + (date.getMonth() + 1)
            : date.getMonth() + 1;
        let monthFolder =
          SelectedDepartmentText +
          "/" +
          project +
          "/Email/" +
          monthNumber +
          "-" +
          date.getFullYear();
        var emlurl = GetEMLUrl(strFileName, monthFolder);
  
        UploadeMailOnOneDrive(MIMEData, emlurl);
        console.log(CurentSelectedEmail);
        let selectedInsertId;
        let emlResponse = EmailResult;
        if (!!emlResponse) {
          debugger;
          // let eml = Office.context.mailbox.item;
          let outlookMail = {
            ToEmail:
              CurentSelectedEmail.toRecipients.length > 0
                ? CurentSelectedEmail.toRecipients[0].emailAddress.address + ";"
                : "", //eml.to[0].emailAddress,
            FromEmail: CurentSelectedEmail.from.emailAddress.address, //  eml.from.emailAddress,
            CC:
              CurentSelectedEmail.ccRecipients.length > 0
                ? CurentSelectedEmail.ccRecipients
                    .map((item) => item.emailAddress["address"])
                    .join(";")
                : [], //eml.cc,
            BCC:
              CurentSelectedEmail.bccRecipients.length > 0
                ? CurentSelectedEmail.bccRecipients
                    .map((item) => item.emailAddress["address"])
                    .join(";")
                : [], // eml.bcc,
            Subject: CurentSelectedEmail.subject, // eml.subject,
            EmailDate: new Date(CurentSelectedEmail.createdDateTime), //  eml.dateTimeCreated,
            EmailBody: CurentSelectedEmail.body["content"],
            IsAttachment: CurentSelectedEmail.hasAttachments, // eml.attachments.length > 0 ? true : false,
            TagId: parseInt(TagId),
            CreatedBy: MyEmail,
            ModifiedBy: MyEmail,
            MessageID: MessageId,
            Department: SelectedDepartmentText,
            EMLOneDriveID: emlResponse.id,
            EMLDownloadLink: emlResponse["@content.downloadUrl"],
          };
          // save into database email result
          selectedInsertId = AddEmailMasterGroupWise(outlookMail);
        }
        var folder = SelectedDepartmentText + "/" + project;
        let fileName;
        let attachUrl;
        let attachResponse;
        $("#progessbar").val("40");
        if (mailItem.hasAttachments) {
          let attachment = getSelectedMailAttachmentItem(MessageId);
          console.log(attachment.value[0].name);
          fileName = attachment.value[0].name; //mailItem.attachments[0].name;
          attachUrl = GetAttachmentUrl(fileName, folder);
          attachResponse = UploadAttachmentOnOneDrive(
            attachUrl,
            attachment.value[0]
          );
          //save into databse attachment result
          AddEmailAttachmentGroupwise(attachResponse, fileName, selectedInsertId);
        }
        $("#progessbar").val("70");
        AddAssignCategory("Archived", CurentSelectedEmail, "Archived");
        AddAssignCategory(TagTextValue, CurentSelectedEmail, "Tag");
        CreateAndMoveToFolder(mailItem);
      } else {
        AddAssignCategory("Archived", CurentSelectedEmail, "Archived");
        AddAssignCategory(TagTextValue, CurentSelectedEmail, "Tag");
      }
    }
  }
  function UploadeMailOnOneDrive(body, url) {
    let response;
    if (body.length < 4000000) {
      $.ajax({
        url: url,
        async: false,
        type: "PUT",
        contentType: "text/html",
        crossDomain: true,
        data: body,
        headers: { Authorization: "Bearer " + AccessToken },
        success: function (results) {
          console.log(results);
          response = results;
        },
        error: function (errorRes) {
          console.log("Error in getting data: " + errorRes);
          if (errorRes.status == 401) {
            requestToken();
            UploadeMailOnOneDrive(body, url);
          }
        },
      });
      //  GetAccessTokenFromRefreshToken();
    } else {
      console.log("body length greater than 4MB");
      url = url.replace("/content", "/createUploadSession");
      $.ajax({
        url: url,
        async: false,
        type: "POST",
        crossDomain: true,
        headers: { Authorization: "Bearer " + AccessToken },
        success: function (results) {
          console.log(results);
          let url = results.uploadUrl;
          UploadFileBySession(url, body);
        },
        error: function (errorRes) {
          console.log("Error in getting data: " + errorRes);
        },
      });
    }
    return response;
  }
  function UploadFileBySession(url, body) {
    // Step 2: Read the file or obtain a file reference
    // Replace this with your own file reading logic
    const fileData = body; // File contents or file reference
    let response;
    // Step 3: Upload the file in chunks
    const chunkSize = 5 * 1024 * 1024; // Chunk size (5 MB in this example)
    let start = 0;
    let end = Math.min(chunkSize, fileData.length);
  
    while (start < fileData.length) {
      const chunk = fileData.slice(start, end);
      $.ajax({
        url: url,
        async: false,
        type: "PUT",
        contentType: "text/html",
        crossDomain: true,
        data: chunk,
        headers: {
          "Content-Range": `bytes ${start}-${end - 1}/${fileData.length}`,
          Authorization: "Bearer " + AccessToken,
        },
        success: function (results) {
          console.log(results);
          if (!!results.webUrl) {
            EmailResult = results;
          }
        },
        error: function (errorRes) {
          console.log("Error in getting data: " + errorRes);
        },
      });
      start = end;
      end = Math.min(start + chunkSize, fileData.length);
    }
  }
  function UploadAttachmentOnOneDrive(url, file) {
    debugger;
    //let attdata = Office.context.mailbox.item.attachments;
    let res;
    $.ajax({
      url: url,
      async: false,
      type: "PUT",
      contentType: "text/html",
      crossDomain: true,
      data: file,
      headers: { Authorization: "Bearer " + AccessToken },
      success: function (results) {
        console.log(results);
        res = results;
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
        res = [];
      },
    });
    return res;
  }
  function GetEMLUrl(fileName, folder) {
    let url = "";
    try {
      url = `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}/${fileName}:/content`;
    } catch (e) {
      console.log(e);
    }
    return url;
  }
  function GetAttachmentUrl(fileName, folder) {
    let url = "";
    try {
      url = `https://graph.microsoft.com/v1.0/me/drive/root:/${folder}/Attachments/${fileName}:/content`;
    } catch (e) {
      console.log(e);
    }
    return url;
  }
  //get email id
  const getEmail = () => {
    if (
      Office.context.mailbox.userProfile.emailAddress == "spdata@grunley.com" ||
      Office.context.mailbox.userProfile.emailAddress ==
        "hitendra.b.solanki@tretainfotech.com"
    ) {
      return "shrinandbakshi@grunley.com";
    } else {
      return Office.context.mailbox.userProfile.emailAddress;
    }
  };
  function getSelectedEmails() {
    debugger;
    // Retrieve the subject line of the selected messages and log it to a list in the task pane.
    Office.context.mailbox.getSelectedItemsAsync((asyncResult) => {
      debugger;
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return [];
      }
      return asyncResult.value;
    });
  }
  function GetAccessTokenFromRefreshToken() {
    // let RequestURL = "https://login.microsoftonline.com/common/oauth2/token";
    let RequestURL =
      "https://login.microsoftonline.com/0a6bce9c-73e9-4d31-88fc-240e3069abfe/oauth2/v2.0/authorize";
    let ContentType = "application/x-www-form-urlencoded";
    let RedirectURL = "https://localhost:3000/";
    let OneDriveClientID = "7b8ee4b7-4c4f-4e6b-88a7-b8f30fe849e2";
    let SecretKey = "0Muvy2qb6[.Mh?fRQ[ErKQCrwtpF0GUV";
    let resource = "https://graph.microsoft.com/";
    const accessTokenOption = {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    };
    Office.auth.getAccessToken(accessTokenOption, function (result) {
      debugger;
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(
          `Token retrieval failed with message: ${result.error.message}`
        );
      } else {
        // Use the outlook access token.
        // outlookToken = result.value;
        console.log(result.value);
      }
    });
  }
  function CreateFolder() {
    debugger;
    $.ajax({
      url: "https://outlook.office.com/api/v2.0/me/mailfolders/inbox",
      type: "GET",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + outlookToken },
      success: function (results) {
        console.log(results);
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
  }
  function CreateAndMoveToFolder(mailItem) {
    debugger;
    let messageId = mailItem.id;
    let urlChild = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/childFolders?$filter=displayName eq 'Archived-TIM'`;
    let subFolderName = TagTextValue;
    $.ajax({
      url: urlChild,
      type: "GET",
      async: false,
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      headers: { Authorization: "Bearer " + AccessToken },
      success: function (results) {
        console.log(results);
        let urlsubChild = `https://graph.microsoft.com/v1.0/me/mailFolders/${results.value[0].id}/childFolders?$filter=displayName eq '${subFolderName}'`;
        $.ajax({
          url: urlsubChild,
          type: "GET",
          async: false,
          contentType: "application/json",
          dataType: "json",
          crossDomain: true,
          headers: { Authorization: "Bearer " + AccessToken },
          success: function (childresults) {
            console.log(childresults);
            let subchildFolderResonse = childresults.value;
            if (subchildFolderResonse.length > 0) {
              // let messageId = Office.context.mailbox.item.itemId;
              debugger;
              $.ajax({
                url: `https://graph.microsoft.com/v1.0/me/messages/${messageId}/move`,
                type: "POST",
                contentType: "application/json",
                dataType: "json",
                crossDomain: true,
                async: false,
                headers: { Authorization: "Bearer " + AccessToken },
                data: JSON.stringify({
                  destinationId: subchildFolderResonse[0].id,
                }),
                success: function (movemailresults) {
                  console.log(movemailresults);
                },
                error: function (moveEmailerror) {
                  console.log(
                    "Error in getting data move mail: " + moveEmailerror
                  );
                },
              });
            } else {
              $.ajax({
                url: `https://graph.microsoft.com/v1.0/me/mailFolders/${results.value[0].id}/childFolders`,
                type: "POST",
                contentType: "application/json",
                dataType: "json",
                async: false,
                crossDomain: true,
                headers: { Authorization: "Bearer " + AccessToken },
                data: JSON.stringify({
                  displayName: subFolderName,
                  isHidden: true,
                }),
                success: function (createFolderresults) {
                  console.log(createFolderresults);
                  $.ajax({
                    url: `https://graph.microsoft.com/v1.0/me/messages/${messageId}/move`,
                    type: "POST",
                    contentType: "application/json",
                    dataType: "json",
                    crossDomain: true,
                    async: false,
                    headers: { Authorization: "Bearer " + AccessToken },
                    data: JSON.stringify({
                      destinationId: createFolderresults.id,
                    }),
                    success: function (movemailresults) {
                      console.log(movemailresults);
                    },
                    error: function (moveEmailerror) {
                      console.log(
                        "Error in getting data move mail: " + moveEmailerror
                      );
                    },
                  });
                },
                error: function (createFoldererror) {
                  console.log(
                    "Error in getting data create Folder: " + createFoldererror
                  );
                },
              });
            }
          },
          error: function (childerror) {
            console.log("Error in getting data child : " + childerror);
          },
        });
      },
      error: function (Mainerror) {
        console.log("Error in getting data archive folder: " + Mainerror);
      },
    });
  }
  function AddEmailMasterGroupWise(eml) {
    debugger;
    let responsedata = 0;
    console.log("store Data", eml);
    let requestUrl = ApiUrl + "Addin/AddEmailMaster";
    let requestUrlForOperations = ApiUrl + "Addin/AddOperationDeptEmailResults";
    $.ajax({
      // eml.Department === "Operations" ? requestUrlForOperations :
      url: requestUrl,
      type: "POST",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      data: JSON.stringify(eml),
      success: function (results) {
        console.log(results);
        responsedata = results.Value;
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
    return responsedata;
  }
  function AddEmailAttachmentGroupwise(response, fileName, selectedInsertId) {
    debugger;
    let createddate = new Date();
    let selecteddepartment = $("#Department").val();
    if (selecteddepartment == "Operations") {
      let emailAttachment = {};
      emailAttachment.EmailMasterId = selectedInsertId;
      emailAttachment.OneDriveFileId = response["id"];
      emailAttachment.Name = fileName;
      emailAttachment.CreatedDate = createddate;
      emailAttachment.ViewLink = response["webUrl"];
      emailAttachment.DownloadLink = response["@microsoft.graph.downloadUrl"];
      emailAttachment.DownloadOrginalLink = "";
      emailAttachment.ParentPath = "";
      console.log(emailAttachment);
      AddOperationDeptEmailAttachment(emailAttachment);
    } else {
      let emailAttachment = {};
      emailAttachment.EmailMasterId = selectedInsertId;
      emailAttachment.OneDriveFileId = response["id"];
      emailAttachment.Name = fileName;
      emailAttachment.CreatedDate = createddate;
      emailAttachment.ViewLink = response["webUrl"];
      emailAttachment.DownloadLink = response["@microsoft.graph.downloadUrl"];
      emailAttachment.DownloadOrginalLink = "";
      emailAttachment.ParentPath = "";
      console.log(emailAttachment);
      AddEmailAttachment(emailAttachment);
    }
  }
  function AddEmailAttachment(emailAttachment) {
    let requestUrl = ApiUrl + "Addin/AddEmailAttachment";
    let insertedNewId = -1;
    try {
      let outlookMail = {
        EmailMasterId: emailAttachment.EmailMasterId,
        OneDriveFileId: emailAttachment.OneDriveFileId,
        Name: emailAttachment.Name,
        CreatedDate: emailAttachment.CreatedDate,
        ViewLink: emailAttachment.ViewLink,
        DownloadLink: emailAttachment.DownloadLink,
        DownloadOrginalLink: emailAttachment.DownloadOrginalLink,
        ParentPath: emailAttachment.ParentPath,
      };
      console.log(outlookMail);
      $.ajax({
        url: requestUrl,
        type: "POST",
        contentType: "application/json",
        dataType: "json",
        crossDomain: true,
        async: false,
        data: JSON.stringify(outlookMail),
        success: function (results) {
          console.log(results);
          insertedNewId = results.Value;
        },
        error: function (error) {
          console.log("Error in getting data: " + error);
        },
      });
    } catch (ex) {
      console.log(ex);
    }
    return insertedNewId;
  }
  function AddOperationDeptEmailAttachment(operationDeptEmailAttachment) {
    let requestUrl = ApiUrl + "Addin/AddOperationDeptEmailAttachment";
    let insertedNewId = -1;
    try {
      let outlookMail = {
        EmailMasterId: operationDeptEmailAttachment.EmailMasterId,
        OneDriveFileId: operationDeptEmailAttachment.OneDriveFileId,
        Name: operationDeptEmailAttachment.Name,
        CreatedDate: operationDeptEmailAttachment.CreatedDate,
        ViewLink: operationDeptEmailAttachment.ViewLink,
        DownloadLink: operationDeptEmailAttachment.DownloadLink,
        DownloadOrginalLink: operationDeptEmailAttachment.DownloadOrginalLink,
        ParentPath: operationDeptEmailAttachment.ParentPath,
      };
      console.log(outlookMail);
      $.ajax({
        url: requestUrl,
        type: "POST",
        contentType: "application/json",
        dataType: "json",
        crossDomain: true,
        async: false,
        data: JSON.stringify(outlookMail),
        success: function (results) {
          console.log(results);
          insertedNewId = results.Value;
        },
        error: function (error) {
          console.log("Error in getting data: " + error);
        },
      });
    } catch (ex) {
      console.log(ex);
    }
    return insertedNewId;
  }
  function AddAssignCategory(categoryname, mailItem, type) {
    $("#progessbar").val("80");
    if (!CategoryExists(categoryname, mailItem)) {
      if (type == "Tag") {
        const categoriesToAddTag = [
          {
            displayName: categoryname,
            color: Office.MailboxEnums.CategoryColor.Preset3,
          },
        ];
        Office.context.mailbox.item.categories.addAsync(
          categoriesToAddTag,
          function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Successfully added tag categories");
            } else {
              console.log(
                "categories.addAsync call failed with error: " +
                  asyncResult.error.message
              );
            }
          }
        );
      } else if (type == "Archived") {
        const categoriesToAddArchived = [
          {
            displayName: categoryname,
            color: Office.MailboxEnums.CategoryColor.Preset19,
          },
        ];
        Office.context.mailbox.item.categories.addAsync(
          categoriesToAddArchived,
          function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Successfully added Archived categories");
            } else {
              console.log(
                "categories.addAsync call failed with error: " +
                  asyncResult.error.message
              );
            }
          }
        );
      } else {
        const categoriesToAddCat = [
          {
            displayName: categoryname,
            color: Office.MailboxEnums.CategoryColor.Preset22,
          },
        ];
        Office.context.mailbox.item.categories.addAsync(
          categoriesToAddCat,
          function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Successfully added categories");
            } else {
              console.log(
                "categories.addAsync call failed with error: " +
                  asyncResult.error.message
              );
            }
          }
        );
      }
    } else {
      let res = AddCategory(mailItem, categoryname);
      if (!!res) {
        $("#progessbar").val("100");
      }
    }
  }
  function AddCategory(mailItem, categoryname) {
    let index;
    console.log(mailItem);
    let urlData = `https://graph.microsoft.com/v1.0/me/messages/'${mailItem.id}'`;
    $.ajax({
      url: urlData,
      type: "PATCH",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      data: JSON.stringify({
        categories: [categoryname],
      }),
      success: function (categoryResult) {
        console.log(categoryResult);
        index = categoryResult;
      },
      error: function (error) {
        console.log("Error in getting data: category Result " + error);
      },
    });
    return index;
  }
  async function CategoryExists(categoryName, mailItem) {
    debugger;
    let category = false;
    const categoriesToAdd = [categoryName];
    console.log(mailItem);
    // Office.context.mailbox.masterCategories.getAsync(
    //   categoriesToAdd,
    //   function (asyncResult) {
    //     if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    //       console.log("Action failed with error: " + asyncResult.error.message);
    //     } else {
    //       let data = asyncResult.value;
    //       let filterdata = data.filter(
    //         (item) => item.displayName == categoryName
    //       );
    //       return filterdata.length > 0 ? (category = true) : false;
    //     }
    //   }
    // );
    $.ajax({
      url: "https://graph.microsoft.com/v1.0/me/outlook/masterCategories/",
      type: "GET",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      success: function (categoryResult) {
        let data = categoryResult;
        let filterdata = data.filter((item) => item.displayName == categoryName);
        filterdata.length > 0 ? (category = true) : false;
      },
      error: function (error) {
        console.log("Error in getting data: category Result " + error);
      },
    });
    return category;
  }
  function CheckIfTagAssigned(Department, TagId, MessageId, FromEmail) {
    debugger;
    let result = false;
    let checkTagAssigne = {
      Department: Department,
      TagId: TagId,
      MessageId: MessageId,
      FromEmail: FromEmail,
    };
    $.ajax({
      url: ApiUrl + "Addin/CheckIfTagAssigned",
      type: "POST",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      data: JSON.stringify(checkTagAssigne),
      success: function (results) {
        console.log(results);
        result = results.Value;
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
    return result;
  }
  
  $(document).on("click", "#shareLageFile", function (event) {
    btnUploadLargeFile_Click();
  });
  function btnUploadLargeFile_Click() {
    let ProjectNumber = $("#Project").val();
    let ddlDepartment = $("#Department").val();
    let projNumber = ProjectNumber;
    let dialog;
    if (ddlDepartment == "Operations") {
      if (ProjectNumber.split("-").length > 2) {
        projNumber =
          ProjectNumber.split("-")[0] +
          "-" +
          ProjectNumber.split("-")[1] +
          "-" +
          ProjectNumber.split("-")[2];
      }
      projNumber = projNumber.Trim();
    } else {
      projNumber = ddlDepartment;
    }
    if (projNumber != "") {
      projNumber = projNumber.replace(/[*'\",_&#^@]/g, "");
      var response = CreateFolderOnSharePointLibrary(
        projNumber.replace(" ", "").trim()
      );
      console.log(response);
      if (!!response || response != "") {
        if (response.split(";").length > 0) {
          window.open(
            response.split(";")[1] + "/Shared%20Documents/General/Shared Files"
          );
        } else {
          console.log(response);
        }
      } else {
        console.log("Something went wrong, please contact administrator!");
      }
    } else {
      console.log("Please select proper project number/department!");
    }
  }
  function CreateFolderOnSharePointLibrary(projectNumber) {
    let response;
    if (AccessToken != "") {
      let groupID = GetOffice365Group(projectNumber + "@grunley.onmicrosoft.com");
      let siteURL;
      if (groupID != "") {
        siteURL = GetTeamsSiteURL(groupID);
      }
      if (siteURL.split().length > 0) {
        let docLibID = GetSharePointDocumentLibrary(siteURL.split(";")[0]);
        let url = `https://graph.microsoft.com/v1.0/drives/${docLibID}/root:/General/Shared Files`;
        let newFolder = "Shared Files";
        if (
          checkFolderExist(
            `https://graph.microsoft.com/v1.0/drives/${docLibID}/root:/General`,
            "General"
          )
        ) {
          let check = checkFolderExist(url, newFolder);
          debugger;
          if (!check) {
            url = `https://graph.microsoft.com/v1.0/drives/${docLibID}/items/root:/General:/children`;
            CreateOneDriveFolder(url, newFolder);
            response = "Success;" + siteURL.split(";")[1];
          } else {
            response = "Success;" + siteURL.split(";")[1];
          }
        } else {
          response =
            "Teams channel folder is not created, Go to Teams channel and click on Files tab, it will create folder.";
        }
      } else {
        response =
          "Teams Channel or Site does not exist, Please create Teams channel first.";
      }
    }
    return response;
  }
  function GetOffice365Group(DepartmentName) {
    let result;
    if (AccessToken != "") {
      $.ajax({
        url:
          "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified') and mail eq '" +
          DepartmentName +
          "' &$select=id,displayName,description,groupname,groupTypes,mailNickname,mail",
        type: "GET",
        contentType: "application/x-www-form-urlencoded",
        dataType: "json",
        crossDomain: true,
        async: false,
        headers: { Authorization: "Bearer " + AccessToken },
        success: function (results) {
          console.log(results);
          result = results["value"][0]["id"];
        },
        error: function (error) {
          console.log("Error in getting data: " + error);
        },
      });
    }
    return result;
  }
  function GetTeamsSiteURL(GroupID) {
    let result;
    let graphurl =
      "https://graph.microsoft.com/v1.0/groups/" + GroupID + "/sites/root";
    if (AccessToken != "") {
      $.ajax({
        url: graphurl,
        type: "GET",
        contentType: "application/x-www-form-urlencoded",
        dataType: "json",
        crossDomain: true,
        async: false,
        headers: { Authorization: "Bearer " + AccessToken },
        success: function (siteDetails) {
          console.log(siteDetails);
          result = siteDetails["id"] + ";" + siteDetails["webUrl"];
        },
        error: function (error) {
          console.log("Error in getting data: " + error);
        },
      });
    }
    return result;
  }
  function GetSharePointDocumentLibrary(SiteName) {
    let result = "";
    let graphurl =
      "https://graph.microsoft.com/v1.0/sites/" + SiteName + "/drives";
    if (AccessToken != "") {
      $.ajax({
        url: graphurl,
        type: "GET",
        contentType: "application/x-www-form-urlencoded",
        dataType: "json",
        crossDomain: true,
        async: false,
        headers: { Authorization: "Bearer " + AccessToken },
        success: function (siteDetails) {
          console.log(siteDetails);
          result = siteDetails["value"][0]["id"];
        },
        error: function (error) {
          console.log("Error in getting data: " + error);
        },
      });
    }
    return result;
  }
  function checkFolderExist(RequestURL, folderName) {
    let IsFolderExist = false;
    if (AccessToken != "") {
      $.ajax({
        url: RequestURL,
        type: "GET",
        contentType: "application/x-www-form-urlencoded",
        dataType: "json",
        crossDomain: true,
        async: false,
        headers: { Authorization: "Bearer " + AccessToken },
        success: function (projectDetails) {
          console.log(projectDetails);
          if (folderName == projectDetails["name"]) {
            IsFolderExist = true;
          }
        },
        error: function (error) {
          console.log("Error in getting data: " + error);
        },
      });
    }
    return IsFolderExist;
  }
  function CreateOneDriveFolder(RequestUrl, foldername) {
    if (AccessToken != "") {
      $.ajax({
        url: RequestUrl,
        type: "POST",
        contentType: "application/json",
        dataType: "json",
        crossDomain: true,
        async: false,
        data: {
          name: foldername,
          folder: {},
          "@microsoft.graph.conflictBehavior": "replace",
        },
        headers: { Authorization: "Bearer " + AccessToken },
        success: function (projectDetails) {
          console.log(projectDetails);
        },
        error: function (error) {
          console.log("Error in getting data: " + error);
        },
      });
    }
  }
  $(document).on("keypress", "#searchText", function (event) {
    let txtSearch = event.target.value;
    if (txtSearch.length < 3) {
      let errorMsg = $("#errMsg");
      if ($("#errMsg")[0].innerText.length == 0) {
        errorMsg.append("minimum 3 letters are required.");
      }
    } else {
      $("#errMsg")[0].innerText = "";
    }
  });
  $(document).on("blur", "#searchText", function (event) {
    let txtSearch = event.target.value;
    if (txtSearch.length < 3) {
      let errorMsg = $("#errMsg");
      if ($("#errMsg")[0].innerText.length == 0) {
        errorMsg.append("minimum 3 letters are required.");
      }
    } else {
      $("#errMsg")[0].innerText = "";
    }
  });
  $(document).on("click", "#btnsearchText", function (event) {
    let txtSearch = $("#searchText").val();
    if (txtSearch.length > 0) {
      console.log(txtSearch);
      let fromEmail = Office.context.mailbox.item.sender.emailAddress;
      console.log(fromEmail);
      let selecteddepartment = $("#Department").val();
      let selectedTag = $("#Project").val();
      if (
        txtSearch.length > 2 &&
        selectedTag.length > 0 &&
        selecteddepartment.length > 0
      ) {
        if (selecteddepartment == "Operations") {
          let data = GetSearchOperationDeptEmailResult(
            txtSearch,
            selectedTag,
            fromEmail
          );
          console.log(data);
        } else {
          GetSearchEmailResult(
            txtSearch,
            selectedTag,
            selecteddepartment,
            fromEmail
          );
        }
      } else {
        $("#errMsg")[0].innerHTML = "<p>Please Select Department and Project</p>";
      }
    } else {
      $("#errMsg")[0].innerHTML = "<p>minimum 3 letters are required.</p>";
    }
  });
  function GetSearchOperationDeptEmailResult(SearchText, TagId, FromEmail) {
    let requestUrl = `${ApiUrl}/addin/GetSearchOperationDeptEmailResult?SearchText=${SearchText}&TagId=${TagId}&From=${FromEmail}`;
    $.ajax({
      url: requestUrl,
      type: "GET",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      success: function (results) {
        console.log(results);
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
  }
  function GetSearchEmailResult(SearchText, TagId, Department, FromEmail) {
    let requestUrl =
      ApiUrl +
      "/addin/GetSearchEmailResult?SearchText=" +
      SearchText +
      "&TagId=" +
      TagId +
      "&Department=" +
      Department +
      "&From=" +
      FromEmail;
    $.ajax({
      url: requestUrl,
      type: "GET",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      success: function (results) {
        console.log(results);
        BindDate(results);
        FilterData = results;
        $("#errMsg")[0].innerText = "";
        if (results.length == 0) {
          $("#errMsg")[0].innerText = "Error, no emails found!";
        }
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
  }
  $(document).on("click", "#chkArchive", function (event) {
    if (event.target.checked) {
      let Email = getEmail();
      let ArchiveData = FilterData.filter((item) => item.CreatedBy == Email);
      console.log(ArchiveData);
      BindDate(ArchiveData);
    } else {
      BindDate(FilterData);
    }
  });
  function BindDate(results) {
    $("#itemData").empty();
    results.map((item, index) => {
      var htmlData = `<li class="ms-ListItem ms-ListItem--document" tabindex=${index}>
      <span class="ms-ListItem-primaryText">${item.FromEmail}</span> 
     <span class="ms-ListItem-secondaryText">${item.Subject}</span> 
     <span class="ms-ListItem-metaText">${item.EmailDate}${item.EmailTime}</span> 
     <div class="ms-ListItem-selectionTarget"></div>
     <div class="ms-ListItem-actions">
     <div class="ms-ListItem-action">
     <i class="material-icons" data-items=${encodeURIComponent(
       JSON.stringify(item)
     )}  style="color:black">&#xe169;</i>
     
     </div>
     </div>
     </div></li>`;
      $("#itemData").append(htmlData);
    });
  }
  // MessageID=${item.MessageID} Id=${item.ID} EMLOnedriveID=${item.EMLOnedriveID}
  $(document).on("click", ".material-icons", function (event) {
    let data = JSON.parse(decodeURIComponent(event.currentTarget.dataset.items));
    unArchiveToolStripMenuItem_Click(data);
  });
  function unArchiveToolStripMenuItem_Click(data) {
    console.log(data);
    let EmailMasterID = data.ID;
    let MessageID = data.MessageID;
    let EMLOnedriveID = data.EMLOneDriveID;
    let deptsCount;
    let outlookmail = {
      ID: EmailMasterID,
      unArchivedBy: getEmail(),
      Department: $("#Department").val(),
    };
    var selectedinsertid = unArchiveEmail(outlookmail);
    let mailItem = GetEmailByInternalId(MessageID);
  
    if (!!mailItem) {
      deptsCount = RemoveCategories(mailItem);
    }
    let url = "https://graph.microsoft.com/v1.0/me/drive/items/" + EMLOnedriveID;
    RemoveEMLFile(url, "");
    RemoveAttachment(EmailMasterID);
    if (deptsCount == 0 || !!mailItem) {
      MoveFromGrunleyToInbox(mailItem);
    }
    if (selectedinsertid == 1) {
      let NewFilterData = FilterData.filter(
        (item) => item.MessageID != MessageID
      );
      BindDate(NewFilterData);
    }
  }
  function unArchiveEmail(eml) {
    let requestUrl = ApiUrl + "Addin/unArchiveEmail";
    let insertedNewId = -1;
    $.ajax({
      url: requestUrl,
      type: "POST",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      data: JSON.stringify(eml),
      success: function (results) {
        console.log(results);
        insertedNewId = results["Value"];
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
    return insertedNewId;
  }
  function RemoveCategories(mailItem) {
    let index = 0;
    console.log(mailItem);
    let urlData = `https://graph.microsoft.com/v1.0/me/messages/'${mailItem.id}'`;
    $.ajax({
      url: urlData,
      type: "PATCH",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      data: JSON.stringify({
        categories: [],
      }),
      success: function (categoryResult) {
        console.log(categoryResult);
        index = 1;
      },
      error: function (error) {
        console.log("Error in getting data: category Result " + error);
      },
    });
    return index;
  }
  function RemoveEMLFile(url) {
    try {
      deleteFile(url, "");
    } catch (ex) {
      console.log(ex);
    }
  }
  function MoveFromGrunleyToInbox(mailItem) {
    $.ajax({
      url: `https://graph.microsoft.com/v1.0/me/messages/${mailItem.id}/move`,
      type: "POST",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      headers: { Authorization: "Bearer " + AccessToken },
      data: JSON.stringify({
        destinationId: "inbox",
      }),
      success: function (results) {
        console.log(results);
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
  }
  function RemoveAttachment(EmailMasterID) {
    let attachmentIds = GetAttachmentOnedriveID(
      $("#Department").val(),
      EmailMasterID
    );
    for (let index = 0; index < attachmentIds.length; index++) {
      const element = attachmentIds[index];
      let url = "https://graph.microsoft.com/v1.0/me/drive/items/" + element;
      deleteFile(url, "");
    }
  }
  function deleteFile(url, filename) {
    if (AccessToken != "") {
      DeleteFileFromOneDrive(url, filename);
    }
  }
  function DeleteFileFromOneDrive(url, filename) {
    $.ajax({
      url: url,
      type: "DELETE",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      success: function (results) {
        console.log(results);
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
  }
  function GetAttachmentOnedriveID(Department, EmailMasterID) {
    let attachmentOnedriveIDs;
    let requestUrl =
      ApiUrl +
      "/addin/GetAttachmentOnedriveID?Department=" +
      Department +
      "&EmailMasterID=" +
      EmailMasterID;
    let procoreProjectEmails;
    $.ajax({
      url: requestUrl,
      type: "GET",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      success: function (results) {
        console.log(results);
        attachmentOnedriveIDs = results;
      },
      error: function (error) {
        console.log("Error in getting data: " + error);
      },
    });
  
    return attachmentOnedriveIDs;
  }
  function GetEmailByInternalId(ID) {
    let response = null;
    $.ajax({
      url: `https://graph.microsoft.com/v1.0/me/messages?$filter=internetMessageId eq '${ID}'`,
      type: "GET",
      contentType: "application/json",
      dataType: "json",
      crossDomain: true,
      async: false,
      headers: { Authorization: "Bearer " + AccessToken },
      success: function (Result) {
        debugger;
        console.log(Result.value[0]);
        response = Result.value[0];
      },
      error: function (error) {
        console.log("Error in getting data: category Result " + error);
      },
    });
    return response;
  }
  