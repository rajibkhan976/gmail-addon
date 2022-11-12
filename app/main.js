var composeCard = CardService.newCardBuilder();

composeCard.setHeader(
  CardService.newCardHeader()
    .setTitle("Sign your file with origin")
    .setImageStyle(CardService.ImageStyle.CIRCLE)
    .setImageUrl(
      "https://lh3.googleusercontent.com/-PMpkNnjQOOo/YupaP5EyukI/AAAAAAAAAB4/EF6MN1aOQ14m6P8-bdqEglZwaZT4rhztgCNcBGAsYHQ/s400/provenant.png"
    )
);

var composeCardSection =
  CardService.newCardSection().setHeader("<b>Attachments<b/>");
var contextCard = CardService.newCardBuilder();
var contextCardSection = CardService.newCardSection();
var selectedFilesHashArr = [];
var passcodeStr = "";
var drafts = GmailApp.getDrafts();
var draftMessage = drafts[0]?.getMessage();
var attachments = draftMessage?.getAttachments();
var emailID = Session.getActiveUser().getEmail();

var timestamp = Date.now();
var signingDate = new Date();
var dd = String(signingDate.getDate()).padStart(2, "0");
var mm = String(signingDate.getMonth() + 1).padStart(2, "0"); //January is 0!
var yyyy = signingDate.getFullYear();

signingDate = mm + "/" + dd + "/" + yyyy;

var hashStore = PropertiesService.getScriptProperties();
var userSharedFileCache = CacheService.getUserCache();
var credentialCache = CacheService.getScriptCache();

const options = {
  method: "GET",
  contentType: "application/json",
  muteHttpExceptions: true,
};

var response = UrlFetchApp.fetch(
  "https://fbc0-117-205-40-252.in.ngrok.io/api/identifier/5e887ac47eed253091be10cb/credential?type=received",
  options
);

if (response?.getResponseCode() === 200.0) {
  var credentials = response.getContentText();
  var credJson = JSON.parse(credentials);
  credentialCache.put("credential", JSON.stringify(credJson));
}

/**
 * Compose trigger function that fires when the compose UI is
 * requested.
 *
 * @param {event} e The compose trigger event object.
 *
 * @return {Card[]}
 */
function onComposeWindowAddonIconClick(event) {
  return [buildComposeCard(event)];
}

/**
 * Context trigger function that fires when the user opens
 * an email.
 *
 * @param {event} e The context trigger event object.
 *
 * @return {Card[]}
 */
function onGmailMessageOpen(event) {
  return [buildContexCard(event)];
}

/**
 * Homepage trigger function that fires when the user clicks
 * on the addon icon.
 *
 * @param {event} e The homepage trigger event object.
 *
 * @return {Card[]}
 */
function onRightSidebarAddonIconClick(event) {
  return [buildLoggedInUserHomepage(event)];
}

/**
 * Compose card section
 */

function buildComposeCard(event) {
  // Logger.log(event);

  if (drafts.length === 0) {
    composeCardSection.addWidget(
      CardService.newDecoratedText()
        .setText("<b>Please attach your file :)</b>")
        .setStartIcon(
          CardService.newIconImage().setIcon(CardService.Icon.DESCRIPTION)
        )
    );
  } else {
    for (var i = 0; i < attachments.length; i++) {
      var credentialJson = JSON.parse(credentialCache.get("credential"));

      composeCardSection.addWidget(
        CardService.newDecoratedText()
          .setText(attachments[i].getName())
          .setStartIcon(
            CardService.newIconImage().setIconUrl(
              "https://lh3.googleusercontent.com/-B-l8TROOVF0/Yup8fcpaUaI/AAAAAAAAACI/eOBP8uShxSs8LbAhaIujuwhzLET2yVSEACNcBGAsYHQ/s400/file.png"
            )
          )
          .setWrapText(true)
          .setSwitchControl(
            CardService.newSwitch()
              .setFieldName(`fileHash_${i}`)
              .setValue(attachments[i].getHash())
              .setControlType(CardService.SwitchControlType.CHECK_BOX)
              .setOnChangeAction(
                CardService.newAction()
                  .setFunctionName("handleSwitchChange")
                  .setParameters({
                    key: `fileHash_${i}`,
                    hash: attachments[i].getHash(),
                  })
              )
          )
      );
      composeCardSection.addWidget(
        CardService.newDecoratedText()
          .setText((attachments[i].getSize() / 1024).toFixed(2) + " KB")
          .setTopLabel("Size")
      );
    }

    composeCardSection.addWidget(CardService.newDivider());

    composeCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(
          credentialJson ? credentialJson[0]?.sad?.a?.personLegalName : emailID
        )
        .setTopLabel("Signer")
        .setStartIcon(
          CardService.newIconImage().setIcon(CardService.Icon.EMAIL)
        )
    );

    composeCardSection.addWidget(CardService.newDivider());

    composeCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(signingDate)
        .setTopLabel("Signing date")
        .setStartIcon(
          CardService.newIconImage().setIcon(CardService.Icon.CLOCK)
        )
    );

    // composeCardSection.addWidget(CardService.newDivider());

    // composeCardSection.addWidget(
    //   CardService.newTextInput()
    //     .setFieldName("passcode")
    //     .setTitle("Passcode")
    //     .setOnChangeAction(
    //       CardService.newAction().setFunctionName("handlePasscodeChange")
    //     )
    // );

    var signingAction =
      CardService.newAction().setFunctionName("signAttachments");

    composeCardSection.addWidget(
      CardService.newTextButton()
        .setText("SIGN")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor("rgb(98, 2, 238)")
        .setOnClickAction(signingAction)
    );
  }
  return composeCard.addSection(composeCardSection).build();
}

function handleSwitchChange(event) {
  hashStore.setProperty(event.parameters.key, event.parameters.hash);
}

function handlePasscodeChange(event) {
  passcodeStr =
    event.commonEventObject.formInputs["passcode"].stringInputs.value[0];
}

function downloadSignatureFile() {
  var file = DriveApp.createFile(
    `Signature_${timestamp}.txt`,
    JSON.stringify({
      signer_oobi: "string",
      signer_credential_said: "string",
      signatures: [
        {
          fileHash: "string",
          signature: ["string"],
        },
      ],
    })
  );

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  var actionResponse = CardService.newActionResponseBuilder()
    .setOpenLink(
      CardService.newOpenLink()
        .setUrl(file.getDownloadUrl())
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.NOTHING)
    )
    .build();

  var addAttachment = `<div style="display: block; border: 1px solid #F3F5F5; padding: 0.5rem; border-radius: 2px;"><a id="signature" href=\"${file.getDownloadUrl()}\" style="display: block; text-decoration: none;">${file.getName()}</a><div style="display: block;">Verify with <a href=\"https://provenant.net/\">Origin</a> by <a href=\"https://provenant.net/\">Provenant</a></div></div>`;

  var updateDraftActionResponse =
    CardService.newUpdateDraftActionResponseBuilder()
      .setUpdateDraftBodyAction(
        CardService.newUpdateDraftBodyAction()
          .addUpdateContent(
            addAttachment,
            CardService.ContentType.IMMUTABLE_HTML
          )
          .setUpdateType(CardService.UpdateDraftBodyType.INSERT_AT_END)
      )
      .build();

  return updateDraftActionResponse, actionResponse;
}

function signAttachments() {
  // Make a POST request with form data.
  var keyArr = hashStore.getKeys();
  var credentialJson = JSON.parse(credentialCache.get("credential"));
  var data = {
    credential_said: credentialJson ? credentialJson[0]?.sad?.d : "",
  };

  var attachmentsObjArr = [];
  var attachmentsArr = [];
  var addAttachment = "";

  keyArr.forEach((key, index) => {
    attachmentsObjArr.push({ fileHash: hashStore.getProperty(key) });
    attachmentsArr.push(hashStore.getProperty(key));
  });

  Object.assign(data, { attachments: attachmentsObjArr });

  attachments.forEach((file, index) => {
    if (attachmentsArr.includes(file.getHash())) {
      addAttachment += `<div style="display: block; background-color: #E8E8E8; padding: 0.5rem; margin-bottom: 1rem; border-radius: 2px;"><div style="display: block;">${file.getName()}</div><div>Signed with <a href=\"https://provenant.net/\">Origin</a> by <a href=\"https://provenant.net/\">Provenant</a></div></div>`;
    }
  });

  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(data),
    muteHttpExceptions: true,
  };

  Logger.log(options);

  var response = UrlFetchApp.fetch(
    "https://fbc0-117-205-40-252.in.ngrok.io/api/identifier/5e887ac47eed253091be10cb/sign",
    options
  );

  Logger.log(response);

  if (response?.getResponseCode() === 200.0) {
    var json = response.getContentText();
    var signedData = JSON.parse(json);
    var content = JSON.stringify(signedData);
    hashStore.deleteAllProperties();

    var fileId = userSharedFileCache.get("fileId");

    if (fileId) {
      var file = DriveApp.getFileById(fileId);
      file.setContent(content);
    } else {
      var file = DriveApp.createFile(`Signature_${timestamp}.sig`, content);
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.EDIT
      );
      userSharedFileCache.put("fileId", file.getId(), 300);
    }

    addAttachment += `<div style="display: block; border: 1px solid #F3F5F5; padding: 0.5rem; border-radius: 2px;"><a id="signature" href=\"${file.getDownloadUrl()}\" style="display: block; text-decoration: none;">${file.getName()}</a><div style="display: block;">Verify with <a href=\"https://provenant.net/\">Origin</a> by <a href=\"https://provenant.net/\">Provenant</a></div><div style="display: none;">fileId={${file.getId()}}</div></div>`;

    var updateDraftActionResponse =
      CardService.newUpdateDraftActionResponseBuilder()
        .setUpdateDraftBodyAction(
          CardService.newUpdateDraftBodyAction()
            .addUpdateContent(
              addAttachment,
              CardService.ContentType.IMMUTABLE_HTML
            )
            .setUpdateType(CardService.UpdateDraftBodyType.INSERT_AT_END)
        )
        .build();
  }

  if (response?.getResponseCode() === 200.0 && signedData) {
    return updateDraftActionResponse;
  } else {
    composeCardSection.addWidget(
      CardService.newDecoratedText().setText(
        "<b><font color='#FF0000'>Signing failed!</font></b>"
      )
    );
    return composeCard.addSection(composeCardSection).build();
  }
}

/**
 * Context card section
 */

function buildContexCard(event) {
  // Logger.log(event);
  // Activate temporary Gmail scopes, in this case to allow
  // the add-on to read message metadata and content.
  var accessToken = event.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);

  // Read message metadata and content. This requires the Gmail scope
  // https://www.googleapis.com/auth/gmail.addons.current.message.readonly.
  var messageId = event.gmail.messageId;
  var message = GmailApp.getMessageById(messageId);
  var subject = message.getSubject();
  var sender = message.getFrom();
  var body = message.getPlainBody();
  var messageDate = message.getDate();
  var emailAttachments = message.getAttachments();

  var fileId = body.match(/^fileId={[\w\W]+}$/gim);

  Logger.log(fileId);

  if (emailAttachments.length > 0 && fileId) {
    var id = fileId[0].replace("fileId={", "").replace("}", "");
    var signatureFile = DriveApp.getFileById(id);
    var signatureFileBlob = signatureFile.getBlob();
    var signatureFileContent = signatureFileBlob.getDataAsString();
    var contentArr = signatureFileContent.split(", ");
    var contentObj = {};
    contentArr.forEach((element, index) => {
      var key = element.substring(0, element.indexOf(":"));
      var value = element.substring(element.indexOf(": ") + 1, element.length);
      Object.assign(contentObj, { [key]: value });
    });
  }

  if (
    emailAttachments.length > 0 &&
    !fileId &&
    !emailAttachments.some(
      (element) =>
        element
          .getName()
          .substring(element.getName().length - 4, element.getName().length) ===
        ".sig"
    )
  ) {
    var attachmentsNameArr = [];

    emailAttachments.forEach((item, index) => {
      attachmentsNameArr.push(item.getName());
    });

    contextCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(
          `<b><font color='#6202EE'>Attachment ${attachmentsNameArr.join(
            ", "
          )} is not signed. If the attachment contains higly confidential or financial information you may ask the sender to re-send the attachmnet(s) with a signature so you can be sure of its origin</font></b>`
        )
        .setWrapText(true)
    );

    // contextCardSection.addWidget(
    //   CardService.newDecoratedText()
    //     .setText(`<b>${emailAttachments.length} files found</b>`)
    //     .setTopLabel("in this conversation")
    // );

    // contextCardSection.addWidget(CardService.newDivider());

    // contextCardSection.addWidget(
    //   CardService.newGrid()
    //     .setNumColumns(3)
    //     .addItem(CardService.newGridItem().setIdentifier("itemA").setTitle(" "))
    //     .addItem(
    //       CardService.newGridItem()
    //         .setIdentifier("itemB")
    //         .setImage(
    //           CardService.newImageComponent()
    //             .setImageUrl(
    //               "https://lh3.googleusercontent.com/-IA5V30Kq4r8/YvXjLTSgQLI/AAAAAAAAAC4/prCn2EiFi40cFyxS8_UTubekvruEltFdQCO8EGAYYCw/s400/context.png"
    //             )
    //             .setCropStyle(CardService.newImageCropStyle())
    //             .setBorderStyle(CardService.newBorderStyle())
    //         )
    //     )
    //     .addItem(CardService.newGridItem().setIdentifier("itemC").setTitle(" "))
    // );

    // contextCardSection.addWidget(
    //   CardService.newDecoratedText()
    //     .setText(
    //       "All signed files in the current conversation will show up here."
    //     )
    //     .setWrapText(true)
    // );
  } else {
    contextCardSection.addWidget(
      CardService.newDecoratedText().setText(
        "<b><font color='#6202EE'>No attachments found</font></b>"
      )
    );
  }

  // Logger.log(
  //   emailAttachments.some(
  //     (element) =>
  //       element
  //         .getName()
  //         .substring(element.getName().length - 4, element.getName().length) ===
  //       ".sig"
  //   )
  // );

  if (
    fileId ||
    (emailAttachments.length > 0 &&
      emailAttachments.some(
        (element) =>
          element
            .getName()
            .substring(
              element.getName().length - 4,
              element.getName().length
            ) === ".sig"
      ))
  ) {
    return buildVerifiedSignatureContextCard(event);
  } else {
    return contextCard.addSection(contextCardSection).build();
  }
}

function buildVerifiedSignatureContextCard(event) {
  // Logger.log(event);
  // Activate temporary Gmail scopes, in this case to allow
  // the add-on to read message metadata and content.
  var accessToken = event.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);

  // Read message metadata and content. This requires the Gmail scope
  // https://www.googleapis.com/auth/gmail.addons.current.message.readonly.
  var messageId = event.gmail.messageId;
  var message = GmailApp.getMessageById(messageId);
  var subject = message.getSubject();
  var sender = message.getFrom();
  var body = message.getPlainBody();
  var messageDate = message.getDate();
  var emailAttachments = message.getAttachments();

  var fileId = body.match(/^fileId={[\w\W]+}$/gim);

  var contentObj = null;

  var attachmentsHashArr = [];
  var signedFilesHashArr = [];
  var attachmentsHashObjArr = [];

  var verifiedSignedFileHashArr = [];
  var unverifiedSignedFilesHashArr = [];

  var verifiedSignedFilesNameArr = [];
  var unverifiedSignedFilesNameArr = [];
  var unsignedFilesNameArr = [];

  if (fileId) {
    var id = fileId[0].replace("fileId={", "").replace("}", "");
    var signatureFile = DriveApp.getFileById(id);
    var signatureFileBlob = signatureFile.getBlob();
    var signatureFileContent = signatureFileBlob.getDataAsString();
    contentObj = JSON.parse(signatureFileContent);
  } else if (
    emailAttachments.length > 0 &&
    emailAttachments.some(
      (element) =>
        element
          .getName()
          .substring(element.getName().length - 4, element.getName().length) ===
        ".sig"
    )
  ) {
    var sigFile = emailAttachments.find(
      (element) =>
        element
          .getName()
          .substring(element.getName().length - 4, element.getName().length) ===
        ".sig"
    );
    var signFileContent = sigFile.getDataAsString();
    contentObj = JSON.parse(signFileContent);
  }

  emailAttachments &&
    emailAttachments.length > 0 &&
    emailAttachments.forEach((item, index) => {
      attachmentsHashArr.push(item.getHash());
    });

  contentObj && contentObj.signatures && contentObj.signatures.length > 0;
  contentObj.signatures.forEach((item, index) => {
    if (attachmentsHashArr.includes(item.fileHash)) {
      signedFilesHashArr.push(item.fileHash);
    }
  });

  signedFilesHashArr.length > 0 &&
    signedFilesHashArr.forEach((item, index) => {
      attachmentsHashObjArr.push({ fileHash: item });
    });

  contentObj && attachmentsHashObjArr.length > 0
    ? Object.assign(contentObj, { attachments: attachmentsHashObjArr })
    : (contentObj = null);

  if (!contentObj) {
    emailAttachments.forEach((item) => {
      if (
        item
          .getName()
          .substring(item.getName().length - 4, item.getName().length) !==
        ".sig"
      ) {
        unverifiedSignedFilesNameArr.push(item.getName());
      }
    });
  }

  if (contentObj) {
    const options = {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(contentObj),
      muteHttpExceptions: true,
    };

    Logger.log(options);

    var response = UrlFetchApp.fetch(
      "https://fbc0-117-205-40-252.in.ngrok.io/api/identifier/verify",
      options
    );

    // Logger.log(response);
  }

  if (response?.getResponseCode() === 200.0) {
    var json = response.getContentText();
    var verifiedData = JSON.parse(json);

    verifiedData &&
      verifiedData.signatureVerificationResult &&
      verifiedData.signatureVerificationResult.length > 0 &&
      signedFilesHashArr.length > 0 &&
      verifiedData.signatureVerificationResult.forEach((item, index) => {
        if (
          signedFilesHashArr.includes(item.fileHash) &&
          item.status &&
          item.status.length > 0
        ) {
          item.status.some((element) => element.isValid)
            ? verifiedSignedFileHashArr.push(item.fileHash)
            : unverifiedSignedFilesHashArr.push(item.fileHash);
        }
      });

    emailAttachments.forEach((item, index) => {
      if (verifiedSignedFileHashArr.includes(item.getHash())) {
        verifiedSignedFilesNameArr.push(item.getName());
      }
      if (unverifiedSignedFilesHashArr.includes(item.getHash())) {
        unverifiedSignedFilesNameArr.push(item.getName());
      }
      if (
        !verifiedSignedFileHashArr.includes(item.getHash()) &&
        !unverifiedSignedFilesHashArr.includes(item.getHash()) &&
        item
          .getName()
          .substring(item.getName().length - 4, item.getName().length) !==
          ".sig"
      ) {
        unsignedFilesNameArr.push(item.getName());
      }
    });

    var signingProvenanceAction = CardService.newAction().setFunctionName(
      "buildSigningProvenanceContextCard"
    );

    var verifiedSignatureCardFooter =
      CardService.newFixedFooter().setPrimaryButton(
        CardService.newTextButton()
          .setText("View signing provenance")
          .setOnClickAction(signingProvenanceAction)
      );

    var verifiedSignatureCard = CardService.newCardBuilder().setFixedFooter(
      verifiedSignatureCardFooter
    );

    var verifiedSignatureCardSection = CardService.newCardSection();

    verifiedSignatureCardSection.addWidget(
      CardService.newDecoratedText().setText(
        `<b><font color='#6202EE'>Verified files (${verifiedSignedFilesNameArr.length})</font></b>`
      )
    );

    verifiedSignatureCardSection.addWidget(
      CardService.newImage().setImageUrl(
        `${
          unverifiedSignedFilesNameArr.length === 0
            ? "https://lh3.googleusercontent.com/-a77n7LDJdYQ/YwX8tbFOTgI/AAAAAAAAAEc/cdQaKUwN7e0Zd8tJ3L8lJnaON02jZlNFQCO8EGAYYCw/s400/banner.png"
            : "https://lh3.googleusercontent.com/-YpwtWQD4kDw/YwZPP-ueqXI/AAAAAAAAAFI/e77bx3bRbIUvlz660JjKvRy_Kbg0RikJACNcBGAsYHQ/s400/banner_failed.png"
        }`
      )

      // CardService.newGrid()
      //   .setNumColumns(1)
      //   .addItem(
      //     CardService.newGridItem()
      //       .setIdentifier("itemC")
      //       .setImage(
      //         CardService.newImageComponent()
      //           .setImageUrl(
      //             "https://lh3.googleusercontent.com/-a77n7LDJdYQ/YwX8tbFOTgI/AAAAAAAAAEc/cdQaKUwN7e0Zd8tJ3L8lJnaON02jZlNFQCO8EGAYYCw/s400/banner.png"
      //           )
      //           .setCropStyle(CardService.newImageCropStyle())
      //           .setBorderStyle(CardService.newBorderStyle())
      //       )
      //   )
    );

    verifiedSignedFilesNameArr.length > 0 &&
      verifiedSignedFilesNameArr.forEach((item) => {
        verifiedSignatureCardSection.addWidget(
          CardService.newDecoratedText()
            .setText(
              `<b>Attachment ${item} has a verified signature and a proven origin.</b>`
            )
            .setEndIcon(
              CardService.newIconImage().setIconUrl(
                "https://lh3.googleusercontent.com/-RD_b9WS2zMw/YwX_giWsLNI/AAAAAAAAAEs/NtGwm4WXcV8q44gY5iG3EKh1-iTLcgVEQCO8EGAYYCw/s400/verified.png"
              )
            )
            .setWrapText(true)
        );
      });

    unverifiedSignedFilesNameArr.length > 0 &&
      unverifiedSignedFilesNameArr.forEach((item) => {
        verifiedSignatureCardSection.addWidget(
          CardService.newDecoratedText()
            .setText(
              `<b>Attachment ${item} has a bad signature. This means its origin doesn't appear to match the sender, either because an attacker modified it or the sender is misconfigured. Do not trust it.</b>`
            )
            .setEndIcon(
              CardService.newIconImage().setIconUrl(
                "https://lh3.googleusercontent.com/-rMQxs1i3_iQ/YwZORxIRXGI/AAAAAAAAAE8/dIpc2pZJPNUh2RIQxWMlDf_F4JZnX3buQCO8EGAYYCw/s400/cross.png"
              )
            )
            .setWrapText(true)
        );
      });

    unsignedFilesNameArr.length > 0 &&
      unsignedFilesNameArr.forEach((item) => {
        verifiedSignatureCardSection.addWidget(
          CardService.newDecoratedText()
            .setText(
              `<b><font color='#6202EE'>Attachment ${item} is not signed. If the attachment contains higly confidential or financial information you may ask the sender to re-send the attachmnet(s) with a signature so you can be sure of its origin</font></b>`
            )
            .setEndIcon(
              CardService.newIconImage().setIconUrl(
                "https://lh3.googleusercontent.com/-rMQxs1i3_iQ/YwZORxIRXGI/AAAAAAAAAE8/dIpc2pZJPNUh2RIQxWMlDf_F4JZnX3buQCO8EGAYYCw/s400/cross.png"
              )
            )
            .setWrapText(true)
        );
      });

    var credentialJson = JSON.parse(credentialCache.get("credential"));

    verifiedSignatureCardSection.addWidget(CardService.newDivider());

    verifiedSignatureCardSection.addWidget(
      CardService.newDecoratedText().setText(
        `<b><font color='#6202EE'>Signature details</font></b>`
      )
    );

    verifiedSignatureCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(`<b>${credentialJson[0]?.sad?.a?.personLegalName}</b>`)
        .setTopLabel("Legal Name")
    );

    verifiedSignatureCardSection.addWidget(CardService.newDivider());

    verifiedSignatureCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(`<b>Acme</b>`)
        .setTopLabel("Company Name")
    );

    verifiedSignatureCardSection.addWidget(CardService.newDivider());

    verifiedSignatureCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(`<b>${credentialJson[0]?.sad?.a?.officialRole}</b>`)
        .setTopLabel("Role")
    );

    verifiedSignatureCardSection.addWidget(CardService.newDivider());

    verifiedSignatureCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(
          `<b><font color='#0A4BB8'><u>${credentialJson[0]?.sad?.a?.LEI}</u></font></b>`
        )
        .setTopLabel("LEI")
    );

    verifiedSignatureCardSection.addWidget(CardService.newDivider());
  }

  if (response?.getResponseCode() === 200.0 && verifiedData) {
    return verifiedSignatureCard
      .addSection(verifiedSignatureCardSection)
      .build();
  } else {
    return buildUnverfiedSignatureContextCard(
      event,
      unverifiedSignedFilesNameArr
    );
  }
}

function buildUnverfiedSignatureContextCard(
  event,
  unverifiedSignedFilesNameArr
) {
  // Activate temporary Gmail scopes, in this case to allow
  // the add-on to read message metadata and content.
  var accessToken = event.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);

  // Read message metadata and content. This requires the Gmail scope
  // https://www.googleapis.com/auth/gmail.addons.current.message.readonly.
  var messageId = event.gmail.messageId;
  var message = GmailApp.getMessageById(messageId);
  var subject = message.getSubject();
  var sender = message.getFrom();
  var body = message.getPlainBody();
  var messageDate = message.getDate();
  var emailAttachments = message.getAttachments();

  var unverifiedSignatureCard = CardService.newCardBuilder();
  var unverifiedSignatureCardSection = CardService.newCardSection();

  unverifiedSignatureCardSection.addWidget(
    CardService.newDecoratedText().setText(
      `<b><font color='#6202EE'>Unverified files (${
        unverifiedSignedFilesNameArr?.length
          ? unverifiedSignedFilesNameArr?.length
          : 0
      })</font></b>`
    )
  );

  unverifiedSignatureCardSection.addWidget(
    CardService.newImage().setImageUrl(
      "https://lh3.googleusercontent.com/-YpwtWQD4kDw/YwZPP-ueqXI/AAAAAAAAAFI/e77bx3bRbIUvlz660JjKvRy_Kbg0RikJACNcBGAsYHQ/s400/banner_failed.png"
    )

    // CardService.newGrid()
    //   .setNumColumns(1)
    //   .addItem(
    //     CardService.newGridItem()
    //       .setIdentifier("itemC")
    //       .setImage(
    //         CardService.newImageComponent()
    //           .setImageUrl(
    //             "https://lh3.googleusercontent.com/-a77n7LDJdYQ/YwX8tbFOTgI/AAAAAAAAAEc/cdQaKUwN7e0Zd8tJ3L8lJnaON02jZlNFQCO8EGAYYCw/s400/banner.png"
    //           )
    //           .setCropStyle(CardService.newImageCropStyle())
    //           .setBorderStyle(CardService.newBorderStyle())
    //       )
    //   )
  );

  // unverifiedSignatureCardSection.addWidget(
  //   CardService.newDecoratedText()
  //     .setText(`<b>filename1.pdf</b>`)
  //     .setEndIcon(
  //       CardService.newIconImage().setIconUrl(
  //         "https://lh3.googleusercontent.com/-rMQxs1i3_iQ/YwZORxIRXGI/AAAAAAAAAE4/vQ8UqWowKBIS0nH5D_7wCxNglv1Ep80awCNcBGAsYHQ/s400/cross.png"
  //       )
  //     )
  //     .setWrapText(true)
  // );

  unverifiedSignedFilesNameArr &&
    unverifiedSignedFilesNameArr.length > 0 &&
    unverifiedSignedFilesNameArr.forEach((item) => {
      unverifiedSignatureCardSection.addWidget(
        CardService.newDecoratedText()
          .setText(
            `<b>Attachment ${item} has a bad signature. This means its origin doesn't appear to match the sender, either because an attacker modified it or the sender is misconfigured. Do not trust it.</b>`
          )
          .setEndIcon(
            CardService.newIconImage().setIconUrl(
              "https://lh3.googleusercontent.com/-rMQxs1i3_iQ/YwZORxIRXGI/AAAAAAAAAE8/dIpc2pZJPNUh2RIQxWMlDf_F4JZnX3buQCO8EGAYYCw/s400/cross.png"
            )
          )
          .setWrapText(true)
      );
    });

  // unverifiedSignatureCardSection.addWidget(
  //   CardService.newDecoratedText().setText(
  //     `<b><font color='#6202EE'>Signature details</font></b>`
  //   )
  // );

  // unverifiedSignatureCardSection.addWidget(
  //   CardService.newDecoratedText()
  //     .setText(
  //       `The file has an invalid digital signature. Contact the sender before trusting it.`
  //     )
  //     .setWrapText(true)
  // );

  var credentialJson = JSON.parse(credentialCache.get("credential"));

  unverifiedSignatureCardSection.addWidget(CardService.newDivider());

  unverifiedSignatureCardSection.addWidget(
    CardService.newDecoratedText().setText(
      `<b><font color='#6202EE'>Signature details</font></b>`
    )
  );

  unverifiedSignatureCardSection.addWidget(
    CardService.newDecoratedText()
      .setText(`<b>${credentialJson[0]?.sad?.a?.personLegalName}</b>`)
      .setTopLabel("Legal Name")
  );

  unverifiedSignatureCardSection.addWidget(CardService.newDivider());

  unverifiedSignatureCardSection.addWidget(
    CardService.newDecoratedText()
      .setText(`<b>Acme</b>`)
      .setTopLabel("Company Name")
  );

  unverifiedSignatureCardSection.addWidget(CardService.newDivider());

  unverifiedSignatureCardSection.addWidget(
    CardService.newDecoratedText()
      .setText(`<b>${credentialJson[0]?.sad?.a?.officialRole}</b>`)
      .setTopLabel("Role")
  );

  unverifiedSignatureCardSection.addWidget(CardService.newDivider());

  unverifiedSignatureCardSection.addWidget(
    CardService.newDecoratedText()
      .setText(
        `<b><font color='#0A4BB8'><u>${credentialJson[0]?.sad?.a?.LEI}</u></font></b>`
      )
      .setTopLabel("LEI")
  );

  unverifiedSignatureCardSection.addWidget(CardService.newDivider());

  var learnMoreAction = CardService.newAction().setFunctionName("learnMore");

  unverifiedSignatureCardSection.addWidget(
    CardService.newDecoratedText()
      .setText(`<b><font color='#6202EE'>LEARN MORE</font></b>`)
      .setOnClickAction(learnMoreAction)
  );

  unverifiedSignatureCardSection.addWidget(CardService.newDivider());

  return unverifiedSignatureCard
    .addSection(unverifiedSignatureCardSection)
    .build();
}

function buildSigningProvenanceContextCard(event) {
  Logger.log(event);

  var signingProvenanceCard = CardService.newCardBuilder();
  var signingProvenanceCardSection = CardService.newCardSection();

  signingProvenanceCardSection.addWidget(
    CardService.newDecoratedText()
      .setText("<b><font color='#6202EE'>Signing provenance</font></b>")
      .setEndIcon(CardService.newIconImage().setIcon(CardService.Icon.PERSON))
  );

  signingProvenanceCardSection.addWidget(
    CardService.newDecoratedText().setText(`Alice Doe`).setTopLabel("Issued by")
  );

  signingProvenanceCardSection.addWidget(CardService.newDivider());

  signingProvenanceCardSection.addWidget(
    CardService.newDecoratedText()
      .setText(`ACME Company`)
      .setTopLabel("Issued by")
  );

  signingProvenanceCardSection.addWidget(CardService.newDivider());

  signingProvenanceCardSection.addWidget(
    CardService.newDecoratedText().setText(`QVI, Inc`).setTopLabel("Issued by")
  );

  signingProvenanceCardSection.addWidget(CardService.newDivider());

  signingProvenanceCardSection.addWidget(
    CardService.newDecoratedText().setText(`GLIEF`).setTopLabel("Issued by")
  );

  signingProvenanceCardSection.addWidget(CardService.newDivider());

  return signingProvenanceCard.addSection(signingProvenanceCardSection).build();
}

function learnMore() {}

/**
 * Homepage card section
 */

function buildAddonHomePage(event) {
  var addonHomepageCard = CardService.newCardBuilder();
  var addonHomepageCardSection = CardService.newCardSection();

  addonHomepageCardSection.addWidget(
    CardService.newGrid()
      .setNumColumns(3)
      .addItem(CardService.newGridItem().setIdentifier("itemA").setTitle(" "))
      .addItem(
        CardService.newGridItem()
          .setIdentifier("itemB")
          .setImage(
            CardService.newImageComponent()
              .setImageUrl(
                "https://lh3.googleusercontent.com/-PZbEVYE5p6Y/YwW5r6QPsjI/AAAAAAAAAEM/eHXaUK78nJshiD646gygdWJxoZoxF229wCO8EGAYYCw/s400/logo.png"
              )
              .setCropStyle(CardService.newImageCropStyle())
              .setBorderStyle(CardService.newBorderStyle())
          )
      )
      .addItem(CardService.newGridItem().setIdentifier("itemC").setTitle(" "))
  );

  addonHomepageCardSection.addWidget(
    CardService.newDecoratedText()
      .setText(
        "No more phishing attacks in spam, smishing in text messages, or robocalls on the telephone."
      )
      .setWrapText(true)
  );

  addonHomepageCardSection.addWidget(CardService.newDivider());

  var loginAction = CardService.newAction().setFunctionName(
    "buildLoggedInUserHomepage"
  );

  addonHomepageCardSection.addWidget(
    CardService.newTextButton().setText("LOGIN").setOnClickAction(loginAction)
  );

  var createAccountAction = CardService.newAction().setFunctionName(
    "buildCreateAccountHomepage"
  );

  addonHomepageCardSection.addWidget(
    CardService.newTextButton()
      .setText("CREATE AN ACCOUNT")
      .setOnClickAction(createAccountAction)
  );

  return addonHomepageCard.addSection(addonHomepageCardSection).build();
}

function buildLoggedInUserHomepage(event) {
  // Logger.log(event);

  var credentialJson = JSON.parse(credentialCache.get("credential"));

  if (credentialJson) {
    var loggedInCard = CardService.newCardBuilder();
    var loggedInCardSection = CardService.newCardSection();

    loggedInCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(`<b>${credentialJson[0]?.sad?.a?.personLegalName}</b>`)
        .setBottomLabel(emailID)
        .setStartIcon(
          CardService.newIconImage().setIconUrl(
            "https://lh3.googleusercontent.com/-XFuAr3Eq6Fg/Yxhq_0GWMII/AAAAAAAAAFs/PAm6TE7RkYc1v8VB3czpEF319P3lJVFDQCNcBGAsYHQ/s400/UserIcon.png"
          )
        )
    );

    loggedInCardSection.addWidget(CardService.newDivider());

    loggedInCardSection.addWidget(
      CardService.newDecoratedText().setText(
        `<b><font color='#6202EE'>Signatory details</font></b>`
      )
    );

    loggedInCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(`<b>${credentialJson[0]?.sad?.a?.personLegalName}</b>`)
        .setTopLabel("Legal Name")
    );

    loggedInCardSection.addWidget(CardService.newDivider());

    loggedInCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(`<b>Acme</b>`)
        .setTopLabel("Company Name")
    );

    loggedInCardSection.addWidget(CardService.newDivider());

    loggedInCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(`<b>${credentialJson[0]?.sad?.a?.officialRole}</b>`)
        .setTopLabel("Role")
    );

    loggedInCardSection.addWidget(CardService.newDivider());

    loggedInCardSection.addWidget(
      CardService.newDecoratedText()
        .setText(
          `<b><font color='#0A4BB8'><u>${credentialJson[0]?.sad?.a?.LEI}</u></font></b>`
        )
        .setTopLabel("LEI")
    );

    loggedInCardSection.addWidget(CardService.newDivider());
  }

  if (credentialJson) {
    return loggedInCard.addSection(loggedInCardSection).build();
  } else {
    return buildAddonHomePage(event);
  }
}

function buildCreateAccountHomepage(event) {
  return [buildLoggedInUserHomepage(event)];
}
