"""
Property reference docs:

 - https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/mapping-canonical-property-names-to-mapi-names#tagged-properties
 - https://interoperability.blob.core.windows.net/files/MS-OXPROPS/[MS-OXPROPS].pdf
 - https://fossies.org/linux/libpst/xml/MAPI_definitions.pdf
 - https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/mapi-constants


+----------------+----------------+-------------------------------------------------------------------------------+
| Range minimum  | Range maximum  |                                 Description                                   |
+----------------+----------------+-------------------------------------------------------------------------------+
| 0x0001         | 0x0BFF         | Message object envelope property; reserved                                    |
| 0x0C00         | 0x0DFF         | Recipient property; reserved                                                  |
| 0x0E00         | 0x0FFF         | Non-transmittable Message property; reserved                                  |
| 0x1000         | 0x2FFF         | Message content property; reserved                                            |
| 0x3000         | 0x33FF         | Multi-purpose property that can appear on all or most objects; reserved       |
| 0x3400         | 0x35FF         | Message store property; reserved                                              |
| 0x3600         | 0x36FF         | Folder and address book container property; reserved                          |
| 0x3700         | 0x38FF         | Attachment property; reserved                                                 |
| 0x3900         | 0x39FF         | Address Book object property; reserved                                        |
| 0x3A00         | 0x3BFF         | Mail user object property; reserved                                           |
| 0x3C00         | 0x3CFF         | Distribution list property; reserved                                          |
| 0x3D00         | 0x3DFF         | Profile section property; reserved                                            |
| 0x3E00         | 0x3EFF         | Status object property; reserved                                              |
| 0x4000         | 0x57FF         | Transport-defined envelope property                                           |
| 0x5800         | 0x5FFF         | Transport-defined recipient property                                          |
| 0x6000         | 0x65FF         | User-defined non-transmittable property                                       |
| 0x6600         | 0x67FF         | Provider-defined internal non-transmittable property                          |
| 0x6800         | 0x7BFF         | Message class-defined content property                                        |
| 0x7C00         | 0x7FFF         | Message class-defined non-transmittable property                              |
| 0x8000         | 0xFFFF         | Reserved for mapping to named properties. The exceptions to this rule are     |
|                |                |   some of the address book tagged properties (those with names beginning with |
|                |                |   PIDTagAddressBook). Many are static property IDs but are in this range.     |
+----------------+----------------+-------------------------------------------------------------------------------+
"""  # noqa: E501
PidTagAccess = 0x0FF4
PidTagAccessControlListData = 0x3FE0
PidTagAccessLevel = 0x0FF7
PidTagAccount = 0x3A00
PidTagAdditionalRenEntryIds = 0x36D8
PidTagAdditionalRenEntryIdsEx = 0x36D9
PidTagAddressBookAuthorizedSenders = 0x8CD8
PidTagAddressBookContainerId = 0xFFFD
PidTagAddressBookDeliveryContentLength = 0x806A
PidTagAddressBookDisplayNamePrintable = 0x39FF
PidTagAddressBookDisplayTypeExtended = 0x8C93
PidTagAddressBookDistributionListExternalMemberCount = 0x8CE3
PidTagAddressBookDistributionListMemberCount = 0x8CE2
PidTagAddressBookDistributionListMemberSubmitAccepted = 0x8073
PidTagAddressBookDistributionListMemberSubmitRejected = 0x8CDA
PidTagAddressBookDistributionListRejectMessagesFromDLMembers = 0x8CDB
PidTagAddressBookEntryId = 0x663B
PidTagAddressBookExtensionAttribute1 = 0x802D
PidTagAddressBookExtensionAttribute10 = 0x8036
PidTagAddressBookExtensionAttribute11 = 0x8C57
PidTagAddressBookExtensionAttribute12 = 0x8C58
PidTagAddressBookExtensionAttribute13 = 0x8C59
PidTagAddressBookExtensionAttribute14 = 0x8C60
PidTagAddressBookExtensionAttribute15 = 0x8C61
PidTagAddressBookExtensionAttribute2 = 0x802E
PidTagAddressBookExtensionAttribute3 = 0x802F
PidTagAddressBookExtensionAttribute4 = 0x8030
PidTagAddressBookExtensionAttribute5 = 0x8031
PidTagAddressBookExtensionAttribute6 = 0x8032
PidTagAddressBookExtensionAttribute7 = 0x8033
PidTagAddressBookExtensionAttribute8 = 0x8034
PidTagAddressBookExtensionAttribute9 = 0x8035
PidTagAddressBookFolderPathname = 0x8004
PidTagAddressBookHierarchicalChildDepartments = 0x8C9A
PidTagAddressBookHierarchicalDepartmentMembers = 0x8C97
PidTagAddressBookHierarchicalIsHierarchicalGroup = 0x8CDD
PidTagAddressBookHierarchicalParentDepartment = 0x8C99
PidTagAddressBookHierarchicalRootDepartment = 0x8C98
PidTagAddressBookHierarchicalShowInDepartments = 0x8C94
PidTagAddressBookHomeMessageDatabase = 0x8006
PidTagAddressBookIsMaster = 0xFFFB
PidTagAddressBookIsMemberOfDistributionList = 0x8008
PidTagAddressBookManageDistributionList = 0x6704
PidTagAddressBookManager = 0x8005
PidTagAddressBookManagerDistinguishedName = 0x8005
PidTagAddressBookMember = 0x8009
PidTagAddressBookMessageId = 0x674F
PidTagAddressBookModerationEnabled = 0x8CB5
PidTagAddressBookNetworkAddress = 0x8170
PidTagAddressBookObjectDistinguishedName = 0x803C
PidTagAddressBookObjectGuid = 0x8C6D
PidTagAddressBookOrganizationalUnitRootDistinguishedName = 0x8CA8
PidTagAddressBookOwner = 0x800C
PidTagAddressBookOwnerBackLink = 0x8024
PidTagAddressBookParentEntryId = 0xFFFC
PidTagAddressBookPhoneticCompanyName = 0x8C91
PidTagAddressBookPhoneticDepartmentName = 0x8C90
PidTagAddressBookPhoneticDisplayName = 0x8C92
PidTagAddressBookPhoneticGivenName = 0x8C8E
PidTagAddressBookPhoneticSurname = 0x8C8F
PidTagAddressBookProxyAddresses = 0x800F
PidTagAddressBookPublicDelegates = 0x8015
PidTagAddressBookReports = 0x800E
PidTagAddressBookRoomCapacity = 0x0807
PidTagAddressBookRoomContainers = 0x8C96
PidTagAddressBookRoomDescription = 0x0809
PidTagAddressBookSenderHintTranslations = 0x8CAC
PidTagAddressBookSeniorityIndex = 0x8CA0
PidTagAddressBookTargetAddress = 0x8011
PidTagAddressBookUnauthorizedSenders = 0x8CD9
PidTagAddressBookX509Certificate = 0x8C6A
PidTagAddressType = 0x3002
PidTagAlternateRecipientAllowed = 0x0002
PidTagAnr = 0x360C
PidTagArchiveDate = 0x301F
PidTagArchivePeriod = 0x301E
PidTagArchiveTag = 0x3018
PidTagAssistant = 0x3A30
PidTagAssistantTelephoneNumber = 0x3A2E
PidTagAssociated = 0x67AA
PidTagAttachAdditionalInformation = 0x370F
PidTagAttachContentBase = 0x3711
PidTagAttachContentId = 0x3712
PidTagAttachContentLocation = 0x3713
PidTagAttachDataBinary = 0x3701
PidTagAttachDataObject = 0x3701
PidTagAttachEncoding = 0x3702
PidTagAttachExtension = 0x3703
PidTagAttachFilename = 0x3704
PidTagAttachFlags = 0x3714
PidTagAttachLongFilename = 0x3707
PidTagAttachLongPathname = 0x370D
PidTagAttachmentContactPhoto = 0x7FFF
PidTagAttachmentFlags = 0x7FFD
PidTagAttachmentHidden = 0x7FFE
PidTagAttachmentLinkId = 0x7FFA
PidTagAttachMethod = 0x3705
PidTagAttachMimeTag = 0x370E
PidTagAttachNumber = 0x0E21
PidTagAttachPathname = 0x3708
PidTagAttachPayloadClass = 0x371A
PidTagAttachPayloadProviderGuidString = 0x3719
PidTagAttachRendering = 0x3709
PidTagAttachSize = 0x0E20
PidTagAttachTag = 0x370A
PidTagAttachTransportName = 0x370C
PidTagAttributeHidden = 0x10F4
PidTagAttributeReadOnly = 0x10F6
PidTagAutoForwardComment = 0x0004
PidTagAutoForwarded = 0x0005
PidTagAutoResponseSuppress = 0x3FDF
PidTagBirthday = 0x3A42
PidTagBlockStatus = 0x1096
PidTagBody = 0x1000
PidTagBodyContentId = 0x1015
PidTagBodyContentLocation = 0x1014
PidTagBodyHtml = 0x1013
PidTagBusiness2TelephoneNumber = 0x3A1B
PidTagBusiness2TelephoneNumbers = 0x3A1B
PidTagBusinessFaxNumber = 0x3A24
PidTagBusinessHomePage = 0x3A51
PidTagBusinessTelephoneNumber = 0x3A08
PidTagCallbackTelephoneNumber = 0x3A02
PidTagCallId = 0x6806
PidTagCarTelephoneNumber = 0x3A1E
PidTagCdoRecurrenceid = 0x10C5
PidTagChangeKey = 0x65E2
PidTagChangeNumber = 0x67A4
PidTagChildrensNames = 0x3A58
PidTagClientActions = 0x6645
PidTagClientSubmitTime = 0x0039
PidTagCodePageId = 0x66C3
PidTagComment = 0x3004
PidTagCompanyMainTelephoneNumber = 0x3A57
PidTagCompanyName = 0x3A16
PidTagComputerNetworkName = 0x3A49
PidTagConflictEntryId = 0x3FF0
PidTagContainerClass = 0x3613
PidTagContainerContents = 0x360F
PidTagContainerFlags = 0x3600
PidTagContainerHierarchy = 0x360E
PidTagContentCount = 0x3602
PidTagContentFilterSpamConfidenceLevel = 0x4076
PidTagContentUnreadCount = 0x3603
PidTagConversationId = 0x3013
PidTagConversationIndex = 0x0071
PidTagConversationIndexTracking = 0x3016
PidTagConversationTopic = 0x0070
PidTagCountry = 0x3A26
PidTagCreationTime = 0x3007
PidTagCreatorEntryId = 0x3FF9
PidTagCreatorName = 0x3FF8
PidTagCustomerId = 0x3A4A
PidTagDamBackPatched = 0x6647
PidTagDamOriginalEntryId = 0x6646
PidTagDefaultPostMessageClass = 0x36E5
PidTagDeferredActionMessageOriginalEntryId = 0x6741
PidTagDeferredDeliveryTime = 0x000F
PidTagDeferredSendNumber = 0x3FEB
PidTagDeferredSendTime = 0x3FEF
PidTagDeferredSendUnits = 0x3FEC
PidTagDelegatedByRule = 0x3FE3
PidTagDelegateFlags = 0x686B
PidTagDeleteAfterSubmit = 0x0E01
PidTagDeletedCountTotal = 0x670B
PidTagDeletedOn = 0x668F
PidTagDeliverTime = 0x0010
PidTagDepartmentName = 0x3A18
PidTagDepth = 0x3005
PidTagDisplayBcc = 0x0E02
PidTagDisplayCc = 0x0E03
PidTagDisplayName = 0x3001
PidTagDisplayNamePrefix = 0x3A45
PidTagDisplayTo = 0x0E04
PidTagDisplayType = 0x3900
PidTagDisplayTypeEx = 0x3905
PidTagEmailAddress = 0x3003
PidTagEndDate = 0x0061
PidTagEntryId = 0x0FFF
PidTagExceptionEndTime = 0x7FFC
PidTagExceptionReplaceTime = 0x7FF9
PidTagExceptionStartTime = 0x7FFB
PidTagExchangeNTSecurityDescriptor = 0x0E84
PidTagExpiryNumber = 0x3FED
PidTagExpiryTime = 0x0015
PidTagExpiryUnits = 0x3FEE
PidTagExtendedFolderFlags = 0x36DA
PidTagExtendedRuleMessageActions = 0x0E99
PidTagExtendedRuleMessageCondition = 0x0E9A
PidTagExtendedRuleSizeLimit = 0x0E9B
PidTagFaxNumberOfPages = 0x6804
PidTagFlagCompleteTime = 0x1091
PidTagFlagStatus = 0x1090
PidTagFlatUrlName = 0x670E
PidTagFolderAssociatedContents = 0x3610
PidTagFolderId = 0x6748
PidTagFolderType = 0x3601
PidTagFollowupIcon = 0x1095
PidTagFreeBusyCountMonths = 0x6869
PidTagFreeBusyEntryIds = 0x36E4
PidTagFreeBusyMessageEmailAddress = 0x6849
PidTagFreeBusyPublishEnd = 0x6848
PidTagFreeBusyPublishStart = 0x6847
PidTagFreeBusyRangeTimestamp = 0x6868
PidTagFtpSite = 0x3A4C
PidTagGatewayNeedsToRefresh = 0x6846
PidTagGender = 0x3A4D
PidTagGeneration = 0x3A05
PidTagGivenName = 0x3A06
PidTagGovernmentIdNumber = 0x3A07
PidTagHasAttachments = 0x0E1B
PidTagHasDeferredActionMessages = 0x3FEA
PidTagHasNamedProperties = 0x664A
PidTagHasRules = 0x663A
PidTagHierarchyChangeNumber = 0x663E
PidTagHobbies = 0x3A43
PidTagHome2TelephoneNumber = 0x3A2F
PidTagHome2TelephoneNumbers = 0x3A2F
PidTagHomeAddressCity = 0x3A59
PidTagHomeAddressCountry = 0x3A5A
PidTagHomeAddressPostalCode = 0x3A5B
PidTagHomeAddressPostOfficeBox = 0x3A5E
PidTagHomeAddressStateOrProvince = 0x3A5C
PidTagHomeAddressStreet = 0x3A5D
PidTagHomeFaxNumber = 0x3A25
PidTagHomeTelephoneNumber = 0x3A09
PidTagHtml = 0x1013
PidTagICalendarEndTime = 0x10C4
PidTagICalendarReminderNextTime = 0x10CA
PidTagICalendarStartTime = 0x10C3
PidTagIconIndex = 0x1080
PidTagImportance = 0x0017
PidTagInConflict = 0x666C
PidTagInitialDetailsPane = 0x3F08
PidTagInitials = 0x3A0A
PidTagInReplyToId = 0x1042
PidTagInstanceKey = 0x0FF6
PidTagInstanceNum = 0x674E
PidTagInstID = 0x674D
PidTagInternetCodepage = 0x3FDE
PidTagInternetMailOverrideFormat = 0x5902
PidTagInternetMessageId = 0x1035
PidTagInternetReferences = 0x1039
PidTagIpmAppointmentEntryId = 0x36D0
PidTagIpmContactEntryId = 0x36D1
PidTagIpmDraftsEntryId = 0x36D7
PidTagIpmJournalEntryId = 0x36D2
PidTagIpmNoteEntryId = 0x36D3
PidTagIpmTaskEntryId = 0x36D4
PidTagIsdnNumber = 0x3A2D
PidTagJunkAddRecipientsToSafeSendersList = 0x6103
PidTagJunkIncludeContacts = 0x6100
PidTagJunkPermanentlyDelete = 0x6102
PidTagJunkPhishingEnableLinks = 0x6107
PidTagJunkThreshold = 0x6101
PidTagKeyword = 0x3A0B
PidTagLanguage = 0x3A0C
PidTagLastModificationTime = 0x3008
PidTagLastModifierEntryId = 0x3FFB
PidTagLastModifierName = 0x3FFA
PidTagLastVerbExecuted = 0x1081
PidTagLastVerbExecutionTime = 0x1082
PidTagListHelp = 0x1043
PidTagListSubscribe = 0x1044
PidTagListUnsubscribe = 0x1045
PidTagLocalCommitTime = 0x6709
PidTagLocalCommitTimeMax = 0x670A
PidTagLocaleId = 0x66A1
PidTagLocality = 0x3A27
PidTagLocation = 0x3A0D
PidTagMailboxOwnerEntryId = 0x661B
PidTagMailboxOwnerName = 0x661C
PidTagManagerName = 0x3A4E
PidTagMappingSignature = 0x0FF8
PidTagMaximumSubmitMessageSize = 0x666D
PidTagMemberId = 0x6671
PidTagMemberName = 0x6672
PidTagMemberRights = 0x6673
PidTagMessageAttachments = 0x0E13
PidTagMessageCcMe = 0x0058
PidTagMessageClass = 0x001A
PidTagMessageCodepage = 0x3FFD
PidTagMessageDeliveryTime = 0x0E06
PidTagMessageEditorFormat = 0x5909
PidTagMessageFlags = 0x0E07
PidTagMessageHandlingSystemCommonName = 0x3A0F
PidTagMessageLocaleId = 0x3FF1
PidTagMessageRecipientMe = 0x0059
PidTagMessageRecipients = 0x0E12
PidTagMessageSize = 0x0E08
PidTagMessageSizeExtended = 0x0E08
PidTagMessageStatus = 0x0E17
PidTagMessageSubmissionId = 0x0047
PidTagMessageToMe = 0x0057
PidTagMid = 0x674A
PidTagMiddleName = 0x3A44
PidTagMimeSkeleton = 0x64F0
PidTagMobileTelephoneNumber = 0x3A1C
PidTagNativeBody = 0x1016
PidTagNextSendAcct = 0x0E29
PidTagNickname = 0x3A4F
PidTagNonDeliveryReportDiagCode = 0x0C05
PidTagNonDeliveryReportReasonCode = 0x0C04
PidTagNonDeliveryReportStatusCode = 0x0C20
PidTagNonReceiptNotificationRequested = 0x0C06
PidTagNormalizedSubject = 0x0E1D
PidTagObjectType = 0x0FFE
PidTagOfficeLocation = 0x3A19
PidTagOfflineAddressBookContainerGuid = 0x6802
PidTagOfflineAddressBookDistinguishedName = 0x6804
PidTagOfflineAddressBookMessageClass = 0x6803
PidTagOfflineAddressBookName = 0x6800
PidTagOfflineAddressBookSequence = 0x6801
PidTagOfflineAddressBookTruncatedProperties = 0x6805
PidTagOrdinalMost = 0x36E2
PidTagOrganizationalIdNumber = 0x3A10
PidTagOriginalAuthorEntryId = 0x004C
PidTagOriginalAuthorName = 0x004D
PidTagOriginalDeliveryTime = 0x0055
PidTagOriginalDisplayBcc = 0x0072
PidTagOriginalDisplayCc = 0x0073
PidTagOriginalDisplayTo = 0x0074
PidTagOriginalEntryId = 0x3A12
PidTagOriginalMessageClass = 0x004B
PidTagOriginalMessageId = 0x1046
PidTagOriginalSenderAddressType = 0x0066
PidTagOriginalSenderEmailAddress = 0x0067
PidTagOriginalSenderEntryId = 0x005B
PidTagOriginalSenderName = 0x005A
PidTagOriginalSenderSearchKey = 0x005C
PidTagOriginalSensitivity = 0x002E
PidTagOriginalSentRepresentingAddressType = 0x0068
PidTagOriginalSentRepresentingEmailAddress = 0x0069
PidTagOriginalSentRepresentingEntryId = 0x005E
PidTagOriginalSentRepresentingName = 0x005D
PidTagOriginalSentRepresentingSearchKey = 0x005F
PidTagOriginalSubject = 0x0049
PidTagOriginalSubmitTime = 0x004E
PidTagOriginatorDeliveryReportRequested = 0x0023
PidTagOriginatorNonDeliveryReportRequested = 0x0C08
PidTagOscSyncEnabled = 0x7C24
PidTagOtherAddressCity = 0x3A5F
PidTagOtherAddressCountry = 0x3A60
PidTagOtherAddressPostalCode = 0x3A61
PidTagOtherAddressPostOfficeBox = 0x3A64
PidTagOtherAddressStateOrProvince = 0x3A62
PidTagOtherAddressStreet = 0x3A63
PidTagOtherTelephoneNumber = 0x3A1F
PidTagOutOfOfficeState = 0x661D
PidTagOwnerAppointmentId = 0x0062
PidTagPagerTelephoneNumber = 0x3A21
PidTagParentEntryId = 0x0E09
PidTagParentFolderId = 0x6749
PidTagParentKey = 0x0025
PidTagParentSourceKey = 0x65E1
PidTagPersonalHomePage = 0x3A50
PidTagPolicyTag = 0x3019
PidTagPostalAddress = 0x3A15
PidTagPostalCode = 0x3A2A
PidTagPostOfficeBox = 0x3A2B
PidTagPredecessorChangeList = 0x65E3
PidTagPrimaryFaxNumber = 0x3A23
PidTagPrimarySendAccount = 0x0E28
PidTagPrimaryTelephoneNumber = 0x3A1A
PidTagPriority = 0x0026
PidTagProcessed = 0x7D01
PidTagProfession = 0x3A46
PidTagProhibitReceiveQuota = 0x666A
PidTagProhibitSendQuota = 0x666E
PidTagPurportedSenderDomain = 0x4083
PidTagRadioTelephoneNumber = 0x3A1D
PidTagRead = 0x0E69
PidTagReadReceiptAddressType = 0x4029
PidTagReadReceiptEmailAddress = 0x402A
PidTagReadReceiptEntryId = 0x0046
PidTagReadReceiptName = 0x402B
PidTagReadReceiptRequested = 0x0029
PidTagReadReceiptSearchKey = 0x0053
PidTagReadReceiptSmtpAddress = 0x5D05
PidTagReceiptTime = 0x002A
PidTagReceivedByAddressType = 0x0075
PidTagReceivedByEmailAddress = 0x0076
PidTagReceivedByEntryId = 0x003F
PidTagReceivedByName = 0x0040
PidTagReceivedBySearchKey = 0x0051
PidTagReceivedBySmtpAddress = 0x5D07
PidTagReceivedRepresentingAddressType = 0x0077
PidTagReceivedRepresentingEmailAddress = 0x0078
PidTagReceivedRepresentingEntryId = 0x0043
PidTagReceivedRepresentingName = 0x0044
PidTagReceivedRepresentingSearchKey = 0x0052
PidTagReceivedRepresentingSmtpAddress = 0x5D08
PidTagRecipientDisplayName = 0x5FF6
PidTagRecipientEntryId = 0x5FF7
PidTagRecipientFlags = 0x5FFD
PidTagRecipientOrder = 0x5FDF
PidTagRecipientProposed = 0x5FE1
PidTagRecipientProposedEndTime = 0x5FE4
PidTagRecipientProposedStartTime = 0x5FE3
PidTagRecipientReassignmentProhibited = 0x002B
PidTagRecipientTrackStatus = 0x5FFF
PidTagRecipientTrackStatusTime = 0x5FFB
PidTagRecipientType = 0x0C15
PidTagRecordKey = 0x0FF9
PidTagReferredByName = 0x3A47
PidTagRemindersOnlineEntryId = 0x36D5
PidTagRemoteMessageTransferAgent = 0x0C21
PidTagRenderingPosition = 0x370B
PidTagReplyRecipientEntries = 0x004F
PidTagReplyRecipientNames = 0x0050
PidTagReplyRequested = 0x0C17
PidTagReplyTemplateId = 0x65C2
PidTagReplyTime = 0x0030
PidTagReportDisposition = 0x0080
PidTagReportDispositionMode = 0x0081
PidTagReportEntryId = 0x0045
PidTagReportingMessageTransferAgent = 0x6820
PidTagReportName = 0x003A
PidTagReportSearchKey = 0x0054
PidTagReportTag = 0x0031
PidTagReportText = 0x1001
PidTagReportTime = 0x0032
PidTagResolveMethod = 0x3FE7
PidTagResponseRequested = 0x0063
PidTagResponsibility = 0x0E0F
PidTagRetentionDate = 0x301C
PidTagRetentionFlags = 0x301D
PidTagRetentionPeriod = 0x301A
PidTagRights = 0x6639
PidTagRoamingDatatypes = 0x7C06
PidTagRoamingDictionary = 0x7C07
PidTagRoamingXmlStream = 0x7C08
PidTagRowid = 0x3000
PidTagRowType = 0x0FF5
PidTagRtfCompressed = 0x1009
PidTagRtfInSync = 0x0E1F
PidTagRuleActionNumber = 0x6650
PidTagRuleActions = 0x6680
PidTagRuleActionType = 0x6649
PidTagRuleCondition = 0x6679
PidTagRuleError = 0x6648
PidTagRuleFolderEntryId = 0x6651
PidTagRuleId = 0x6674
PidTagRuleIds = 0x6675
PidTagRuleLevel = 0x6683
PidTagRuleMessageLevel = 0x65ED
PidTagRuleMessageName = 0x65EC
PidTagRuleMessageProvider = 0x65EB
PidTagRuleMessageProviderData = 0x65EE
PidTagRuleMessageSequence = 0x65F3
PidTagRuleMessageState = 0x65E9
PidTagRuleMessageUserFlags = 0x65EA
PidTagRuleName = 0x6682
PidTagRuleProvider = 0x6681
PidTagRuleProviderData = 0x6684
PidTagRuleSequence = 0x6676
PidTagRuleState = 0x6677
PidTagRuleUserFlags = 0x6678
PidTagRwRulesStream = 0x6802
PidTagScheduleInfoAppointmentTombstone = 0x686A
PidTagScheduleInfoAutoAcceptAppointments = 0x686D
PidTagScheduleInfoDelegateEntryIds = 0x6845
PidTagScheduleInfoDelegateNames = 0x6844
PidTagScheduleInfoDelegateNamesW = 0x684A
PidTagScheduleInfoDelegatorWantsCopy = 0x6842
PidTagScheduleInfoDelegatorWantsInfo = 0x684B
PidTagScheduleInfoDisallowOverlappingAppts = 0x686F
PidTagScheduleInfoDisallowRecurringAppts = 0x686E
PidTagScheduleInfoDontMailDelegates = 0x6843
PidTagScheduleInfoFreeBusy = 0x686C
PidTagScheduleInfoFreeBusyAway = 0x6856
PidTagScheduleInfoFreeBusyBusy = 0x6854
PidTagScheduleInfoFreeBusyMerged = 0x6850
PidTagScheduleInfoFreeBusyTentative = 0x6852
PidTagScheduleInfoMonthsAway = 0x6855
PidTagScheduleInfoMonthsBusy = 0x6853
PidTagScheduleInfoMonthsMerged = 0x684F
PidTagScheduleInfoMonthsTentative = 0x6851
PidTagScheduleInfoResourceType = 0x6841
PidTagSchedulePlusFreeBusyEntryId = 0x6622
PidTagScriptData = 0x0004
PidTagSearchFolderDefinition = 0x6845
PidTagSearchFolderEfpFlags = 0x6848
PidTagSearchFolderExpiration = 0x683A
PidTagSearchFolderId = 0x6842
PidTagSearchFolderLastUsed = 0x6834
PidTagSearchFolderRecreateInfo = 0x6844
PidTagSearchFolderStorageType = 0x6846
PidTagSearchFolderTag = 0x6847
PidTagSearchFolderTemplateId = 0x6841
PidTagSearchKey = 0x300B
PidTagSecurityDescriptorAsXml = 0x0E6A
PidTagSelectable = 0x3609
PidTagSenderAddressType = 0x0C1E
PidTagSenderEmailAddress = 0x0C1F
PidTagSenderEntryId = 0x0C19
PidTagSenderIdStatus = 0x4079
PidTagSenderName = 0x0C1A
PidTagSenderSearchKey = 0x0C1D
PidTagSenderSmtpAddress = 0x5D01
PidTagSenderTelephoneNumber = 0x6802
PidTagSendInternetEncoding = 0x3A71
PidTagSendRichInfo = 0x3A40
PidTagSensitivity = 0x0036
PidTagSentMailSvrEID = 0x6740
PidTagSentRepresentingAddressType = 0x0064
PidTagSentRepresentingEmailAddress = 0x0065
PidTagSentRepresentingEntryId = 0x0041
PidTagSentRepresentingFlags = 0x401A
PidTagSentRepresentingName = 0x0042
PidTagSentRepresentingSearchKey = 0x003B
PidTagSentRepresentingSmtpAddress = 0x5D02
PidTagSmtpAddress = 0x39FE
PidTagSortLocaleId = 0x6705
PidTagSourceKey = 0x65E0
PidTagSpokenName = 0x8CC2
PidTagSpouseName = 0x3A48
PidTagStartDate = 0x0060
PidTagStartDateEtc = 0x301B
PidTagStateOrProvince = 0x3A28
PidTagStoreEntryId = 0x0FFB
PidTagStoreState = 0x340E
PidTagStoreSupportMask = 0x340D
PidTagStreetAddress = 0x3A29
PidTagSubfolders = 0x360A
PidTagSubject = 0x0037
PidTagSubjectPrefix = 0x003D
PidTagSupplementaryInfo = 0x0C1B
PidTagSurname = 0x3A11
PidTagSwappedToDoData = 0x0E2D
PidTagSwappedToDoStore = 0x0E2C
PidTagTargetEntryId = 0x3010
PidTagTelecommunicationsDeviceForDeafTelephoneNumber = 0x3A4B
PidTagTelexNumber = 0x3A2C
PidTagTemplateData = 0x0001
PidTagTemplateid = 0x3902
PidTagTextAttachmentCharset = 0x371B
PidTagThumbnailPhoto = 0x8C9E
PidTagTitle = 0x3A17
PidTagTnefCorrelationKey = 0x007F
PidTagToDoItemFlags = 0x0E2B
PidTagTransmittableDisplayName = 0x3A20
PidTagTransportMessageHeaders = 0x007D
PidTagTrustSender = 0x0E79
PidTagUserCertificate = 0x3A22
PidTagUserEntryId = 0x6619
PidTagUserX509Certificate = 0x3A70
PidTagViewDescriptorBinary = 0x7001
PidTagViewDescriptorName = 0x7006
PidTagViewDescriptorStrings = 0x7002
PidTagViewDescriptorVersion = 0x7007
PidTagVoiceMessageAttachmentOrder = 0x6805
PidTagVoiceMessageDuration = 0x6801
PidTagVoiceMessageSenderName = 0x6803
PidTagWeddingAnniversary = 0x3A41
PidTagWlinkAddressBookEID = 0x6854
PidTagWlinkAddressBookStoreEID = 0x6891
PidTagWlinkCalendarColor = 0x6853
PidTagWlinkClientID = 0x6890
PidTagWlinkEntryId = 0x684C
PidTagWlinkFlags = 0x684A
PidTagWlinkFolderType = 0x684F
PidTagWlinkGroupClsid = 0x6850
PidTagWlinkGroupHeaderID = 0x6842
PidTagWlinkGroupName = 0x6851
PidTagWlinkOrdinal = 0x684B
PidTagWlinkRecordKey = 0x684D
PidTagWlinkROGroupType = 0x6892
PidTagWlinkSaveStamp = 0x6847
PidTagWlinkSection = 0x6852
PidTagWlinkStoreEntryId = 0x684E
PidTagWlinkType = 0x6849


MAPI_ACKNOWLEDGEMENT_MODE = 0x0001
MAPI_ALTERNATE_RECIPIENT_ALLOWED = 0x0002
MAPI_AUTHORIZING_USERS = 0x0003
MAPI_AUTO_FORWARD_COMMENT = 0x0004
MAPI_AUTO_FORWARDED = 0x0005
MAPI_CONTENT_CONFIDENTIALITY_ALGORITHM_ID = 0x0006
MAPI_CONTENT_CORRELATOR = 0x0007
MAPI_CONTENT_IDENTIFIER = 0x0008
MAPI_CONTENT_LENGTH = 0x0009
MAPI_CONTENT_RETURN_REQUESTED = 0x000A
MAPI_CONVERSATION_KEY = 0x000B
MAPI_CONVERSION_EITS = 0x000C
MAPI_CONVERSION_WITH_LOSS_PROHIBITED = 0x000D
MAPI_CONVERTED_EITS = 0x000E
MAPI_DEFERRED_DELIVERY_TIME = 0x000F
MAPI_DELIVER_TIME = 0x0010
MAPI_DISCARD_REASON = 0x0011
MAPI_DISCLOSURE_OF_RECIPIENTS = 0x0012
MAPI_DL_EXPANSION_HISTORY = 0x0013
MAPI_DL_EXPANSION_PROHIBITED = 0x0014
MAPI_EXPIRY_TIME = 0x0015
MAPI_IMPLICIT_CONVERSION_PROHIBITED = 0x0016
MAPI_IMPORTANCE = 0x0017
MAPI_IPM_ID = 0x0018
MAPI_LATEST_DELIVERY_TIME = 0x0019
MAPI_MESSAGE_CLASS = 0x001A
MAPI_MESSAGE_DELIVERY_ID = 0x001B
MAPI_MESSAGE_SECURITY_LABEL = 0x001E
MAPI_OBSOLETED_IPMS = 0x001F
MAPI_ORIGINALLY_INTENDED_RECIPIENT_NAME = 0x0020
MAPI_ORIGINAL_EITS = 0x0021
MAPI_ORIGINATOR_CERTIFICATE = 0x0022
MAPI_ORIGINATOR_DELIVERY_REPORT_REQUESTED = 0x0023
MAPI_ORIGINATOR_RETURN_ADDRESS = 0x0024
MAPI_PARENT_KEY = 0x0025
MAPI_PRIORITY = 0x0026
MAPI_ORIGIN_CHECK = 0x0027
MAPI_PROOF_OF_SUBMISSION_REQUESTED = 0x0028
MAPI_READ_RECEIPT_REQUESTED = 0x0029
MAPI_RECEIPT_TIME = 0x002A
MAPI_RECIPIENT_REASSIGNMENT_PROHIBITED = 0x002B
MAPI_REDIRECTION_HISTORY = 0x002C
MAPI_RELATED_IPMS = 0x002D
MAPI_ORIGINAL_SENSITIVITY = 0x002E
MAPI_LANGUAGES = 0x002F
MAPI_REPLY_TIME = 0x0030
MAPI_REPORT_TAG = 0x0031
MAPI_REPORT_TIME = 0x0032
MAPI_RETURNED_IPM = 0x0033
MAPI_SECURITY = 0x0034
MAPI_INCOMPLETE_COPY = 0x0035
MAPI_SENSITIVITY = 0x0036
MAPI_SUBJECT = 0x0037
MAPI_SUBJECT_IPM = 0x0038
MAPI_CLIENT_SUBMIT_TIME = 0x0039
MAPI_REPORT_NAME = 0x003A
MAPI_SENT_REPRESENTING_SEARCH_KEY = 0x003B
MAPI_X400_CONTENT_TYPE = 0x003C
MAPI_SUBJECT_PREFIX = 0x003D
MAPI_NON_RECEIPT_REASON = 0x003E
MAPI_RECEIVED_BY_ENTRYID = 0x003F
MAPI_RECEIVED_BY_NAME = 0x0040
MAPI_SENT_REPRESENTING_ENTRYID = 0x0041
MAPI_SENT_REPRESENTING_NAME = 0x0042
MAPI_RCVD_REPRESENTING_ENTRYID = 0x0043
MAPI_RCVD_REPRESENTING_NAME = 0x0044
MAPI_REPORT_ENTRYID = 0x0045
MAPI_READ_RECEIPT_ENTRYID = 0x0046
MAPI_MESSAGE_SUBMISSION_ID = 0x0047
MAPI_PROVIDER_SUBMIT_TIME = 0x0048
MAPI_ORIGINAL_SUBJECT = 0x0049
MAPI_DISC_VAL = 0x004A
MAPI_ORIG_MESSAGE_CLASS = 0x004B
MAPI_ORIGINAL_AUTHOR_ENTRYID = 0x004C
MAPI_ORIGINAL_AUTHOR_NAME = 0x004D
MAPI_ORIGINAL_SUBMIT_TIME = 0x004E
MAPI_REPLY_RECIPIENT_ENTRIES = 0x004F
MAPI_REPLY_RECIPIENT_NAMES = 0x0050
MAPI_RECEIVED_BY_SEARCH_KEY = 0x0051
MAPI_RCVD_REPRESENTING_SEARCH_KEY = 0x0052
MAPI_READ_RECEIPT_SEARCH_KEY = 0x0053
MAPI_REPORT_SEARCH_KEY = 0x0054
MAPI_ORIGINAL_DELIVERY_TIME = 0x0055
MAPI_ORIGINAL_AUTHOR_SEARCH_KEY = 0x0056
MAPI_MESSAGE_TO_ME = 0x0057
MAPI_MESSAGE_CC_ME = 0x0058
MAPI_MESSAGE_RECIP_ME = 0x0059
MAPI_ORIGINAL_SENDER_NAME = 0x005A
MAPI_ORIGINAL_SENDER_ENTRYID = 0x005B
MAPI_ORIGINAL_SENDER_SEARCH_KEY = 0x005C
MAPI_ORIGINAL_SENT_REPRESENTING_NAME = 0x005D
MAPI_ORIGINAL_SENT_REPRESENTING_ENTRYID = 0x005E
MAPI_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY = 0x005F
MAPI_START_DATE = 0x0060
MAPI_END_DATE = 0x0061
MAPI_OWNER_APPT_ID = 0x0062
MAPI_RESPONSE_REQUESTED = 0x0063
MAPI_SENT_REPRESENTING_ADDRTYPE = 0x0064
MAPI_SENT_REPRESENTING_EMAIL_ADDRESS = 0x0065
MAPI_ORIGINAL_SENDER_ADDRTYPE = 0x0066
MAPI_ORIGINAL_SENDER_EMAIL_ADDRESS = 0x0067
MAPI_ORIGINAL_SENT_REPRESENTING_ADDRTYPE = 0x0068
MAPI_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS = 0x0069
MAPI_CONVERSATION_TOPIC = 0x0070
MAPI_CONVERSATION_INDEX = 0x0071
MAPI_ORIGINAL_DISPLAY_BCC = 0x0072
MAPI_ORIGINAL_DISPLAY_CC = 0x0073
MAPI_ORIGINAL_DISPLAY_TO = 0x0074
MAPI_RECEIVED_BY_ADDRTYPE = 0x0075
MAPI_RECEIVED_BY_EMAIL_ADDRESS = 0x0076
MAPI_RCVD_REPRESENTING_ADDRTYPE = 0x0077
MAPI_RCVD_REPRESENTING_EMAIL_ADDRESS = 0x0078
MAPI_ORIGINAL_AUTHOR_ADDRTYPE = 0x0079
MAPI_ORIGINAL_AUTHOR_EMAIL_ADDRESS = 0x007A
MAPI_ORIGINALLY_INTENDED_RECIP_ADDRTYPE = 0x007B
MAPI_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS = 0x007C
MAPI_TRANSPORT_MESSAGE_HEADERS = 0x007D
MAPI_DELEGATION = 0x007E
MAPI_TNEF_CORRELATION_KEY = 0x007F
MAPI_CONTENT_INTEGRITY_CHECK = 0x0C00
MAPI_EXPLICIT_CONVERSION = 0x0C01
MAPI_IPM_RETURN_REQUESTED = 0x0C02
MAPI_MESSAGE_TOKEN = 0x0C03
MAPI_NDR_REASON_CODE = 0x0C04
MAPI_NDR_DIAG_CODE = 0x0C05
MAPI_NON_RECEIPT_NOTIFICATION_REQUESTED = 0x0C06
MAPI_DELIVERY_POINT = 0x0C07
MAPI_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED = 0x0C08
MAPI_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT = 0x0C09
MAPI_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY = 0x0C0A
MAPI_PHYSICAL_DELIVERY_MODE = 0x0C0B
MAPI_PHYSICAL_DELIVERY_REPORT_REQUEST = 0x0C0C
MAPI_PHYSICAL_FORWARDING_ADDRESS = 0x0C0D
MAPI_PHYSICAL_FORWARDING_ADDRESS_REQUESTED = 0x0C0E
MAPI_PHYSICAL_FORWARDING_PROHIBITED = 0x0C0F
MAPI_PHYSICAL_RENDITION_ATTRIBUTES = 0x0C10
MAPI_PROOF_OF_DELIVERY = 0x0C11
MAPI_PROOF_OF_DELIVERY_REQUESTED = 0x0C12
MAPI_RECIPIENT_CERTIFICATE = 0x0C13
MAPI_RECIPIENT_NUMBER_FOR_ADVICE = 0x0C14
MAPI_RECIPIENT_TYPE = 0x0C15
MAPI_REGISTERED_MAIL_TYPE = 0x0C16
MAPI_REPLY_REQUESTED = 0x0C17
MAPI_REQUESTED_DELIVERY_METHOD = 0x0C18
MAPI_SENDER_ENTRYID = 0x0C19
MAPI_SENDER_NAME = 0x0C1A
MAPI_SUPPLEMENTARY_INFO = 0x0C1B
MAPI_TYPE_OF_MTS_USER = 0x0C1C
MAPI_SENDER_SEARCH_KEY = 0x0C1D
MAPI_SENDER_ADDRTYPE = 0x0C1E
MAPI_SENDER_EMAIL_ADDRESS = 0x0C1F
MAPI_CURRENT_VERSION = 0x0E00
MAPI_DELETE_AFTER_SUBMIT = 0x0E01
MAPI_DISPLAY_BCC = 0x0E02
MAPI_DISPLAY_CC = 0x0E03
MAPI_DISPLAY_TO = 0x0E04
MAPI_PARENT_DISPLAY = 0x0E05
MAPI_MESSAGE_DELIVERY_TIME = 0x0E06
MAPI_MESSAGE_FLAGS = 0x0E07
MAPI_MESSAGE_SIZE = 0x0E08
MAPI_PARENT_ENTRYID = 0x0E09
MAPI_SENTMAIL_ENTRYID = 0x0E0A
MAPI_CORRELATE = 0x0E0C
MAPI_CORRELATE_MTSID = 0x0E0D
MAPI_DISCRETE_VALUES = 0x0E0E
MAPI_RESPONSIBILITY = 0x0E0F
MAPI_SPOOLER_STATUS = 0x0E10
MAPI_TRANSPORT_STATUS = 0x0E11
MAPI_MESSAGE_RECIPIENTS = 0x0E12
MAPI_MESSAGE_ATTACHMENTS = 0x0E13
MAPI_SUBMIT_FLAGS = 0x0E14
MAPI_RECIPIENT_STATUS = 0x0E15
MAPI_TRANSPORT_KEY = 0x0E16
MAPI_MSG_STATUS = 0x0E17
MAPI_MESSAGE_DOWNLOAD_TIME = 0x0E18
MAPI_CREATION_VERSION = 0x0E19
MAPI_MODIFY_VERSION = 0x0E1A
MAPI_HASATTACH = 0x0E1B
MAPI_BODY_CRC = 0x0E1C
MAPI_NORMALIZED_SUBJECT = 0x0E1D
MAPI_RTF_IN_SYNC = 0x0E1F
MAPI_ATTACH_SIZE = 0x0E20
MAPI_ATTACH_NUM = 0x0E21
MAPI_PREPROCESS = 0x0E22
MAPI_ORIGINATING_MTA_CERTIFICATE = 0x0E25
MAPI_PROOF_OF_SUBMISSION = 0x0E26
MAPI_PRIMARY_SEND_ACCOUNT = 0x0E28
MAPI_NEXT_SEND_ACCT = 0x0E29
MAPI_ACCESS = 0x0FF4
MAPI_ROW_TYPE = 0x0FF5
MAPI_INSTANCE_KEY = 0x0FF6
MAPI_ACCESS_LEVEL = 0x0FF7
MAPI_MAPPING_SIGNATURE = 0x0FF8
MAPI_RECORD_KEY = 0x0FF9
MAPI_STORE_RECORD_KEY = 0x0FFA
MAPI_STORE_ENTRYID = 0x0FFB
MAPI_MINI_ICON = 0x0FFC
MAPI_ICON = 0x0FFD
MAPI_OBJECT_TYPE = 0x0FFE
MAPI_ENTRYID = 0x0FFF
MAPI_BODY = 0x1000
MAPI_REPORT_TEXT = 0x1001
MAPI_ORIGINATOR_AND_DL_EXPANSION_HISTORY = 0x1002
MAPI_REPORTING_DL_NAME = 0x1003
MAPI_REPORTING_MTA_CERTIFICATE = 0x1004
MAPI_RTF_SYNC_BODY_CRC = 0x1006
MAPI_RTF_SYNC_BODY_COUNT = 0x1007
MAPI_RTF_SYNC_BODY_TAG = 0x1008
MAPI_RTF_COMPRESSED = 0x1009
MAPI_RTF_SYNC_PREFIX_COUNT = 0x1010
MAPI_RTF_SYNC_TRAILING_COUNT = 0x1011
MAPI_ORIGINALLY_INTENDED_RECIP_ENTRYID = 0x1012
MAPI_BODY_HTML = 0x1013
MAPI_NATIVE_BODY = 0x1016
MAPI_SMTP_MESSAGE_ID = 0x1035
MAPI_INTERNET_REFERENCES = 0x1039
MAPI_IN_REPLY_TO_ID = 0x1042
MAPI_INTERNET_RETURN_PATH = 0x1046
MAPI_ICON_INDEX = 0x1080
MAPI_LAST_VERB_EXECUTED = 0x1081
MAPI_LAST_VERB_EXECUTION_TIME = 0x1082
MAPI_URL_COMP_NAME = 0x10F3
MAPI_ATTRIBUTE_HIDDEN = 0x10F4
MAPI_ATTRIBUTE_SYSTEM = 0x10F5
MAPI_ATTRIBUTE_READ_ONLY = 0x10F6
MAPI_ROWID = 0x3000
MAPI_DISPLAY_NAME = 0x3001
MAPI_ADDRTYPE = 0x3002
MAPI_EMAIL_ADDRESS = 0x3003
MAPI_COMMENT = 0x3004
MAPI_DEPTH = 0x3005
MAPI_PROVIDER_DISPLAY = 0x3006
MAPI_CREATION_TIME = 0x3007
MAPI_LAST_MODIFICATION_TIME = 0x3008
MAPI_RESOURCE_FLAGS = 0x3009
MAPI_PROVIDER_DLL_NAME = 0x300A
MAPI_SEARCH_KEY = 0x300B
MAPI_PROVIDER_UID = 0x300C
MAPI_PROVIDER_ORDINAL = 0x300D
MAPI_TARGET_ENTRY_ID = 0x3010
MAPI_CONVERSATION_ID = 0x3013
MAPI_CONVERSATION_INDEX_TRACKING = 0x3016
MAPI_FORM_VERSION = 0x3301
MAPI_FORM_CLSID = 0x3302
MAPI_FORM_CONTACT_NAME = 0x3303
MAPI_FORM_CATEGORY = 0x3304
MAPI_FORM_CATEGORY_SUB = 0x3305
MAPI_FORM_HOST_MAP = 0x3306
MAPI_FORM_HIDDEN = 0x3307
MAPI_FORM_DESIGNER_NAME = 0x3308
MAPI_FORM_DESIGNER_GUID = 0x3309
MAPI_FORM_MESSAGE_BEHAVIOR = 0x330A
MAPI_DEFAULT_STORE = 0x3400
MAPI_STORE_SUPPORT_MASK = 0x340D
MAPI_STORE_STATE = 0x340E
MAPI_STORE_UNICODE_MASK = 0x340F
MAPI_IPM_SUBTREE_SEARCH_KEY = 0x3410
MAPI_IPM_OUTBOX_SEARCH_KEY = 0x3411
MAPI_IPM_WASTEBASKET_SEARCH_KEY = 0x3412
MAPI_IPM_SENTMAIL_SEARCH_KEY = 0x3413
MAPI_MDB_PROVIDER = 0x3414
MAPI_RECEIVE_FOLDER_SETTINGS = 0x3415
MAPI_VALID_FOLDER_MASK = 0x35DF
MAPI_IPM_SUBTREE_ENTRYID = 0x35E0
MAPI_IPM_OUTBOX_ENTRYID = 0x35E2
MAPI_IPM_WASTEBASKET_ENTRYID = 0x35E3
MAPI_IPM_SENTMAIL_ENTRYID = 0x35E4
MAPI_VIEWS_ENTRYID = 0x35E5
MAPI_COMMON_VIEWS_ENTRYID = 0x35E6
MAPI_FINDER_ENTRYID = 0x35E7
MAPI_CONTAINER_FLAGS = 0x3600
MAPI_FOLDER_TYPE = 0x3601
MAPI_CONTENT_COUNT = 0x3602
MAPI_CONTENT_UNREAD = 0x3603
MAPI_CREATE_TEMPLATES = 0x3604
MAPI_DETAILS_TABLE = 0x3605
MAPI_SEARCH = 0x3607
MAPI_SELECTABLE = 0x3609
MAPI_SUBFOLDERS = 0x360A
MAPI_STATUS = 0x360B
MAPI_ANR = 0x360C
MAPI_CONTENTS_SORT_ORDER = 0x360D
MAPI_CONTAINER_HIERARCHY = 0x360E
MAPI_CONTAINER_CONTENTS = 0x360F
MAPI_FOLDER_ASSOCIATED_CONTENTS = 0x3610
MAPI_DEF_CREATE_DL = 0x3611
MAPI_DEF_CREATE_MAILUSER = 0x3612
MAPI_CONTAINER_CLASS = 0x3613
MAPI_CONTAINER_MODIFY_VERSION = 0x3614
MAPI_AB_PROVIDER_ID = 0x3615
MAPI_DEFAULT_VIEW_ENTRYID = 0x3616
MAPI_ASSOC_CONTENT_COUNT = 0x3617
MAPI_ATTACHMENT_X400_PARAMETERS = 0x3700
MAPI_ATTACH_DATA_OBJ = 0x3701
MAPI_ATTACH_ENCODING = 0x3702
MAPI_ATTACH_EXTENSION = 0x3703
MAPI_ATTACH_FILENAME = 0x3704
MAPI_ATTACH_METHOD = 0x3705
MAPI_ATTACH_LONG_FILENAME = 0x3707
MAPI_ATTACH_PATHNAME = 0x3708
MAPI_ATTACH_RENDERING = 0x3709
MAPI_ATTACH_TAG = 0x370A
MAPI_RENDERING_POSITION = 0x370B
MAPI_ATTACH_TRANSPORT_NAME = 0x370C
MAPI_ATTACH_LONG_PATHNAME = 0x370D
MAPI_ATTACH_MIME_TAG = 0x370E
MAPI_ATTACH_ADDITIONAL_INFO = 0x370F
MAPI_ATTACH_MIME_SEQUENCE = 0x3710
MAPI_ATTACH_CONTENT_ID = 0x3712
MAPI_ATTACH_CONTENT_LOCATION = 0x3713
MAPI_ATTACH_FLAGS = 0x3714
MAPI_DISPLAY_TYPE = 0x3900
MAPI_TEMPLATEID = 0x3902
MAPI_PRIMARY_CAPABILITY = 0x3904
MAPI_SMTP_ADDRESS = 0x39FE
MAPI_7BIT_DISPLAY_NAME = 0x39FF
MAPI_ACCOUNT = 0x3A00
MAPI_ALTERNATE_RECIPIENT = 0x3A01
MAPI_CALLBACK_TELEPHONE_NUMBER = 0x3A02
MAPI_CONVERSION_PROHIBITED = 0x3A03
MAPI_DISCLOSE_RECIPIENTS = 0x3A04
MAPI_GENERATION = 0x3A05
MAPI_GIVEN_NAME = 0x3A06
MAPI_GOVERNMENT_ID_NUMBER = 0x3A07
MAPI_BUSINESS_TELEPHONE_NUMBER = 0x3A08
MAPI_HOME_TELEPHONE_NUMBER = 0x3A09
MAPI_INITIALS = 0x3A0A
MAPI_KEYWORD = 0x3A0B
MAPI_LANGUAGE = 0x3A0C
MAPI_LOCATION = 0x3A0D
MAPI_MAIL_PERMISSION = 0x3A0E
MAPI_MHS_COMMON_NAME = 0x3A0F
MAPI_ORGANIZATIONAL_ID_NUMBER = 0x3A10
MAPI_SURNAME = 0x3A11
MAPI_ORIGINAL_ENTRYID = 0x3A12
MAPI_ORIGINAL_DISPLAY_NAME = 0x3A13
MAPI_ORIGINAL_SEARCH_KEY = 0x3A14
MAPI_POSTAL_ADDRESS = 0x3A15
MAPI_COMPANY_NAME = 0x3A16
MAPI_TITLE = 0x3A17
MAPI_DEPARTMENT_NAME = 0x3A18
MAPI_OFFICE_LOCATION = 0x3A19
MAPI_PRIMARY_TELEPHONE_NUMBER = 0x3A1A
MAPI_BUSINESS2_TELEPHONE_NUMBER = 0x3A1B
MAPI_MOBILE_TELEPHONE_NUMBER = 0x3A1C
MAPI_RADIO_TELEPHONE_NUMBER = 0x3A1D
MAPI_CAR_TELEPHONE_NUMBER = 0x3A1E
MAPI_OTHER_TELEPHONE_NUMBER = 0x3A1F
MAPI_TRANSMITABLE_DISPLAY_NAME = 0x3A20
MAPI_PAGER_TELEPHONE_NUMBER = 0x3A21
MAPI_USER_CERTIFICATE = 0x3A22
MAPI_PRIMARY_FAX_NUMBER = 0x3A23
MAPI_BUSINESS_FAX_NUMBER = 0x3A24
MAPI_HOME_FAX_NUMBER = 0x3A25
MAPI_COUNTRY = 0x3A26
MAPI_LOCALITY = 0x3A27
MAPI_STATE_OR_PROVINCE = 0x3A28
MAPI_STREET_ADDRESS = 0x3A29
MAPI_POSTAL_CODE = 0x3A2A
MAPI_POST_OFFICE_BOX = 0x3A2B
MAPI_TELEX_NUMBER = 0x3A2C
MAPI_ISDN_NUMBER = 0x3A2D
MAPI_ASSISTANT_TELEPHONE_NUMBER = 0x3A2E
MAPI_HOME2_TELEPHONE_NUMBER = 0x3A2F
MAPI_ASSISTANT = 0x3A30
MAPI_SEND_RICH_INFO = 0x3A40
MAPI_WEDDING_ANNIVERSARY = 0x3A41
MAPI_BIRTHDAY = 0x3A42
MAPI_HOBBIES = 0x3A43
MAPI_MIDDLE_NAME = 0x3A44
MAPI_DISPLAY_NAME_PREFIX = 0x3A45
MAPI_PROFESSION = 0x3A46
MAPI_PREFERRED_BY_NAME = 0x3A47
MAPI_SPOUSE_NAME = 0x3A48
MAPI_COMPUTER_NETWORK_NAME = 0x3A49
MAPI_CUSTOMER_ID = 0x3A4A
MAPI_TTYTDD_PHONE_NUMBER = 0x3A4B
MAPI_FTP_SITE = 0x3A4C
MAPI_GENDER = 0x3A4D
MAPI_MANAGER_NAME = 0x3A4E
MAPI_NICKNAME = 0x3A4F
MAPI_PERSONAL_HOME_PAGE = 0x3A50
MAPI_BUSINESS_HOME_PAGE = 0x3A51
MAPI_CONTACT_VERSION = 0x3A52
MAPI_CONTACT_ENTRYIDS = 0x3A53
MAPI_CONTACT_ADDRTYPES = 0x3A54
MAPI_CONTACT_DEFAULT_ADDRESS_INDEX = 0x3A55
MAPI_CONTACT_EMAIL_ADDRESSES = 0x3A56
MAPI_COMPANY_MAIN_PHONE_NUMBER = 0x3A57
MAPI_CHILDRENS_NAMES = 0x3A58
MAPI_HOME_ADDRESS_CITY = 0x3A59
MAPI_HOME_ADDRESS_COUNTRY = 0x3A5A
MAPI_HOME_ADDRESS_POSTAL_CODE = 0x3A5B
MAPI_HOME_ADDRESS_STATE_OR_PROVINCE = 0x3A5C
MAPI_HOME_ADDRESS_STREET = 0x3A5D
MAPI_HOME_ADDRESS_POST_OFFICE_BOX = 0x3A5E
MAPI_OTHER_ADDRESS_CITY = 0x3A5F
MAPI_OTHER_ADDRESS_COUNTRY = 0x3A60
MAPI_OTHER_ADDRESS_POSTAL_CODE = 0x3A61
MAPI_OTHER_ADDRESS_STATE_OR_PROVINCE = 0x3A62
MAPI_OTHER_ADDRESS_STREET = 0x3A63
MAPI_OTHER_ADDRESS_POST_OFFICE_BOX = 0x3A64
MAPI_SEND_INTERNET_ENCODING = 0x3A71
MAPI_STORE_PROVIDERS = 0x3D00
MAPI_AB_PROVIDERS = 0x3D01
MAPI_TRANSPORT_PROVIDERS = 0x3D02
MAPI_DEFAULT_PROFILE = 0x3D04
MAPI_AB_SEARCH_PATH = 0x3D05
MAPI_AB_DEFAULT_DIR = 0x3D06
MAPI_AB_DEFAULT_PAB = 0x3D07
MAPI_FILTERING_HOOKS = 0x3D08
MAPI_SERVICE_NAME = 0x3D09
MAPI_SERVICE_DLL_NAME = 0x3D0A
MAPI_SERVICE_ENTRY_NAME = 0x3D0B
MAPI_SERVICE_UID = 0x3D0C
MAPI_SERVICE_EXTRA_UIDS = 0x3D0D
MAPI_SERVICES = 0x3D0E
MAPI_SERVICE_SUPPORT_FILES = 0x3D0F
MAPI_SERVICE_DELETE_FILES = 0x3D10
MAPI_AB_SEARCH_PATH_UPDATE = 0x3D11
MAPI_PROFILE_NAME = 0x3D12
MAPI_IDENTITY_DISPLAY = 0x3E00
MAPI_IDENTITY_ENTRYID = 0x3E01
MAPI_RESOURCE_METHODS = 0x3E02
MAPI_RESOURCE_TYPE = 0x3E03
MAPI_STATUS_CODE = 0x3E04
MAPI_IDENTITY_SEARCH_KEY = 0x3E05
MAPI_OWN_STORE_ENTRYID = 0x3E06
MAPI_RESOURCE_PATH = 0x3E07
MAPI_STATUS_STRING = 0x3E08
MAPI_X400_DEFERRED_DELIVERY_CANCEL = 0x3E09
MAPI_HEADER_FOLDER_ENTRYID = 0x3E0A
MAPI_REMOTE_PROGRESS = 0x3E0B
MAPI_REMOTE_PROGRESS_TEXT = 0x3E0C
MAPI_REMOTE_VALIDATE_OK = 0x3E0D
MAPI_CONTROL_FLAGS = 0x3F00
MAPI_CONTROL_STRUCTURE = 0x3F01
MAPI_CONTROL_TYPE = 0x3F02
MAPI_DELTAX = 0x3F03
MAPI_DELTAY = 0x3F04
MAPI_XPOS = 0x3F05
MAPI_YPOS = 0x3F06
MAPI_CONTROL_ID = 0x3F07
MAPI_INITIAL_DETAILS_PANE = 0x3F08
MAPI_UNCOMPRESSED_BODY = 0x3FD9
MAPI_INTERNET_CODEPAGE = 0x3FDE
MAPI_AUTO_RESPONSE_SUPPRESS = 0x3FDF
MAPI_MESSAGE_LOCALE_ID = 0x3FF1
MAPI_RULE_TRIGGER_HISTORY = 0x3FF2
MAPI_MOVE_TO_STORE_ENTRYID = 0x3FF3
MAPI_MOVE_TO_FOLDER_ENTRYID = 0x3FF4
MAPI_STORAGE_QUOTA_LIMIT = 0x3FF5
MAPI_EXCESS_STORAGE_USED = 0x3FF6
MAPI_SVR_GENERATING_QUOTA_MSG = 0x3FF7
MAPI_CREATOR_NAME = 0x3FF8
MAPI_CREATOR_ENTRY_ID = 0x3FF9
MAPI_LAST_MODIFIER_NAME = 0x3FFA
MAPI_LAST_MODIFIER_ENTRY_ID = 0x3FFB
MAPI_REPLY_RECIPIENT_SMTP_PROXIES = 0x3FFC
MAPI_MESSAGE_CODEPAGE = 0x3FFD
MAPI_EXTENDED_ACL_DATA = 0x3FFE
MAPI_SENDER_FLAGS = 0x4019
MAPI_SENT_REPRESENTING_FLAGS = 0x401A
MAPI_RECEIVED_BY_FLAGS = 0x401B
MAPI_RECEIVED_REPRESENTING_FLAGS = 0x401C
MAPI_CREATOR_ADDRESS_TYPE = 0x4022
MAPI_CREATOR_EMAIL_ADDRESS = 0x4023
MAPI_SENDER_SIMPLE_DISPLAY_NAME = 0x4030
MAPI_SENT_REPRESENTING_SIMPLE_DISPLAY_NAME = 0x4031
MAPI_RECEIVED_REPRESENTING_SIMPLE_DISPLAY_NAME = 0x4035
MAPI_CREATOR_SIMPLE_DISP_NAME = 0x4038
MAPI_LAST_MODIFIER_SIMPLE_DISPLAY_NAME = 0x4039
MAPI_CONTENT_FILTER_SPAM_CONFIDENCE_LEVEL = 0x4076
MAPI_INTERNET_MAIL_OVERRIDE_FORMAT = 0x5902
MAPI_MESSAGE_EDITOR_FORMAT = 0x5909
MAPI_SENDER_SMTP_ADDRESS = 0x5D01
MAPI_SENT_REPRESENTING_SMTP_ADDRESS = 0x5D02
MAPI_READ_RECEIPT_SMTP_ADDRESS = 0x5D05
MAPI_RECEIVED_BY_SMTP_ADDRESS = 0x5D07
MAPI_RECEIVED_REPRESENTING_SMTP_ADDRESS = 0x5D08
MAPI_SENDING_SMTP_ADDRESS = 0x5D09
MAPI_SIP_ADDRESS = 0x5FE5
MAPI_RECIPIENT_DISPLAY_NAME = 0x5FF6
MAPI_RECIPIENT_ENTRYID = 0x5FF7
MAPI_RECIPIENT_FLAGS = 0x5FFD
MAPI_RECIPIENT_TRACKSTATUS = 0x5FFF
MAPI_CHANGE_KEY = 0x65E2
MAPI_PREDECESSOR_CHANGE_LIST = 0x65E3
MAPI_ID_SECURE_MIN = 0x67F0
MAPI_ID_SECURE_MAX = 0x67FF
MAPI_VOICE_MESSAGE_DURATION = 0x6801
MAPI_SENDER_TELEPHONE_NUMBER = 0x6802
MAPI_VOICE_MESSAGE_SENDER_NAME = 0x6803
MAPI_FAX_NUMBER_OF_PAGES = 0x6804
MAPI_VOICE_MESSAGE_ATTACHMENT_ORDER = 0x6805
MAPI_CALL_ID = 0x6806
MAPI_ATTACHMENT_LINK_ID = 0x7FFA
MAPI_EXCEPTION_START_TIME = 0x7FFB
MAPI_EXCEPTION_END_TIME = 0x7FFC
MAPI_ATTACHMENT_FLAGS = 0x7FFD
MAPI_ATTACHMENT_HIDDEN = 0x7FFE
MAPI_ATTACHMENT_CONTACT_PHOTO = 0x7FFF
MAPI_FILE_UNDER = 0x8005
MAPI_FILE_UNDER_ID = 0x8006
MAPI_CONTACT_ITEM_DATA = 0x8007
MAPI_REFERRED_BY = 0x800E
MAPI_DEPARTMENT = 0x8010
MAPI_HAS_PICTURE = 0x8015
MAPI_HOME_ADDRESS = 0x801A
MAPI_WORK_ADDRESS = 0x801B
MAPI_OTHER_ADDRESS = 0x801C
MAPI_POSTAL_ADDRESS_ID = 0x8022
MAPI_CONTACT_CHARACTER_SET = 0x8023
MAPI_AUTO_LOG = 0x8025
MAPI_FILE_UNDER_LIST = 0x8026
MAPI_EMAIL_LIST = 0x8027
MAPI_ADDRESS_BOOK_PROVIDER_EMAIL_LIST = 0x8028
MAPI_ADDRESS_BOOK_PROVIDER_ARRAY_TYPE = 0x8029
MAPI_HTML = 0x802B
MAPI_YOMI_FIRST_NAME = 0x802C
MAPI_YOMI_LAST_NAME = 0x802D
MAPI_YOMI_COMPANY_NAME = 0x802E
MAPI_BUSINESS_CARD_DISPLAY_DEFINITION = 0x8040
MAPI_BUSINESS_CARD_CARD_PICTURE = 0x8041
MAPI_WORK_ADDRESS_STREET = 0x8045
MAPI_WORK_ADDRESS_CITY = 0x8046
MAPI_WORK_ADDRESS_STATE = 0x8047
MAPI_WORK_ADDRESS_POSTAL_CODE = 0x8048
MAPI_WORK_ADDRESS_COUNTRY = 0x8049
MAPI_WORK_ADDRESS_POST_OFFICE_BOX = 0x804A
MAPI_DISTRIBUTION_LIST_CHECKSUM = 0x804C
MAPI_BIRTHDAY_EVENT_ENTRY_ID = 0x804D
MAPI_ANNIVERSARY_EVENT_ENTRY_ID = 0x804E
MAPI_CONTACT_USER_FIELD1 = 0x804F
MAPI_CONTACT_USER_FIELD2 = 0x8050
MAPI_CONTACT_USER_FIELD3 = 0x8051
MAPI_CONTACT_USER_FIELD4 = 0x8052
MAPI_DISTRIBUTION_LIST_NAME = 0x8053
MAPI_DISTRIBUTION_LIST_ONE_OFF_MEMBERS = 0x8054
MAPI_DISTRIBUTION_LIST_MEMBERS = 0x8055
MAPI_INSTANT_MESSAGING_ADDRESS = 0x8062
MAPI_DISTRIBUTION_LIST_STREAM = 0x8064
MAPI_EMAIL_DISPLAY_NAME = 0x8080
MAPI_EMAIL_ADDR_TYPE = 0x8082
MAPI_EMAIL_EMAIL_ADDRESS = 0x8083
MAPI_EMAIL_ORIGINAL_DISPLAY_NAME = 0x8084
MAPI_EMAIL1ORIGINAL_ENTRY_ID = 0x8085
MAPI_EMAIL1RICH_TEXT_FORMAT = 0x8086
MAPI_EMAIL1EMAIL_TYPE = 0x8087
MAPI_EMAIL2DISPLAY_NAME = 0x8090
MAPI_EMAIL2ENTRY_ID = 0x8091
MAPI_EMAIL2ADDR_TYPE = 0x8092
MAPI_EMAIL2EMAIL_ADDRESS = 0x8093
MAPI_EMAIL2ORIGINAL_DISPLAY_NAME = 0x8094
MAPI_EMAIL2ORIGINAL_ENTRY_ID = 0x8095
MAPI_EMAIL2RICH_TEXT_FORMAT = 0x8096
MAPI_EMAIL3DISPLAY_NAME = 0x80A0
MAPI_EMAIL3ENTRY_ID = 0x80A1
MAPI_EMAIL3ADDR_TYPE = 0x80A2
MAPI_EMAIL3EMAIL_ADDRESS = 0x80A3
MAPI_EMAIL3ORIGINAL_DISPLAY_NAME = 0x80A4
MAPI_EMAIL3ORIGINAL_ENTRY_ID = 0x80A5
MAPI_EMAIL3RICH_TEXT_FORMAT = 0x80A6
MAPI_FAX1ADDRESS_TYPE = 0x80B2
MAPI_FAX1EMAIL_ADDRESS = 0x80B3
MAPI_FAX1ORIGINAL_DISPLAY_NAME = 0x80B4
MAPI_FAX1ORIGINAL_ENTRY_ID = 0x80B5
MAPI_FAX2ADDRESS_TYPE = 0x80C2
MAPI_FAX2EMAIL_ADDRESS = 0x80C3
MAPI_FAX2ORIGINAL_DISPLAY_NAME = 0x80C4
MAPI_FAX2ORIGINAL_ENTRY_ID = 0x80C5
MAPI_FAX3ADDRESS_TYPE = 0x80D2
MAPI_FAX3EMAIL_ADDRESS = 0x80D3
MAPI_FAX3ORIGINAL_DISPLAY_NAME = 0x80D4
MAPI_FAX3ORIGINAL_ENTRY_ID = 0x80D5
MAPI_FREE_BUSY_LOCATION = 0x80D8
MAPI_HOME_ADDRESS_COUNTRY_CODE = 0x80DA
MAPI_WORK_ADDRESS_COUNTRY_CODE = 0x80DB
MAPI_OTHER_ADDRESS_COUNTRY_CODE = 0x80DC
MAPI_ADDRESS_COUNTRY_CODE = 0x80DD
MAPI_BIRTHDAY_LOCAL = 0x80DE
MAPI_WEDDING_ANNIVERSARY_LOCAL = 0x80DF
MAPI_TASK_STATUS = 0x8101
MAPI_TASK_START_DATE = 0x8104
MAPI_TASK_DUE_DATE = 0x8105
MAPI_TASK_ACTUAL_EFFORT = 0x8110
MAPI_TASK_ESTIMATED_EFFORT = 0x8111
MAPI_TASK_FRECUR = 0x8126
MAPI_SEND_MEETING_AS_ICAL = 0x8200
MAPI_APPOINTMENT_SEQUENCE = 0x8201
MAPI_APPOINTMENT_SEQUENCE_TIME = 0x8202
MAPI_APPOINTMENT_LAST_SEQUENCE = 0x8203
MAPI_CHANGE_HIGHLIGHT = 0x8204
MAPI_BUSY_STATUS = 0x8205
MAPI_FEXCEPTIONAL_BODY = 0x8206
MAPI_APPOINTMENT_AUXILIARY_FLAGS = 0x8207
MAPI_OUTLOOK_LOCATION = 0x8208
MAPI_MEETING_WORKSPACE_URL = 0x8209
MAPI_FORWARD_INSTANCE = 0x820A
MAPI_LINKED_TASK_ITEMS = 0x820C
MAPI_APPT_START_WHOLE = 0x820D
MAPI_APPT_END_WHOLE = 0x820E
MAPI_APPOINTMENT_START_TIME = 0x820F
MAPI_APPOINTMENT_END_TIME = 0x8210
MAPI_APPOINTMENT_END_DATE = 0x8211
MAPI_APPOINTMENT_START_DATE = 0x8212
MAPI_APPT_DURATION = 0x8213
MAPI_APPOINTMENT_COLOR = 0x8214
MAPI_APPOINTMENT_SUB_TYPE = 0x8215
MAPI_APPOINTMENT_RECUR = 0x8216
MAPI_APPOINTMENT_STATE_FLAGS = 0x8217
MAPI_RESPONSE_STATUS = 0x8218
MAPI_APPOINTMENT_REPLY_TIME = 0x8220
MAPI_RECURRING = 0x8223
MAPI_INTENDED_BUSY_STATUS = 0x8224
MAPI_APPOINTMENT_UPDATE_TIME = 0x8226
MAPI_EXCEPTION_REPLACE_TIME = 0x8228
MAPI_OWNER_NAME = 0x822E
MAPI_APPOINTMENT_REPLY_NAME = 0x8230
MAPI_RECURRENCE_TYPE = 0x8231
MAPI_RECURRENCE_PATTERN = 0x8232
MAPI_TIME_ZONE_STRUCT = 0x8233
MAPI_TIME_ZONE_DESCRIPTION = 0x8234
MAPI_CLIP_START = 0x8235
MAPI_CLIP_END = 0x8236
MAPI_ORIGINAL_STORE_ENTRY_ID = 0x8237
MAPI_ALL_ATTENDEES_STRING = 0x8238
MAPI_AUTO_FILL_LOCATION = 0x823A
MAPI_TO_ATTENDEES_STRING = 0x823B
MAPI_CCATTENDEES_STRING = 0x823C
MAPI_CONF_CHECK = 0x8240
MAPI_CONFERENCING_TYPE = 0x8241
MAPI_DIRECTORY = 0x8242
MAPI_ORGANIZER_ALIAS = 0x8243
MAPI_AUTO_START_CHECK = 0x8244
MAPI_AUTO_START_WHEN = 0x8245
MAPI_ALLOW_EXTERNAL_CHECK = 0x8246
MAPI_COLLABORATE_DOC = 0x8247
MAPI_NET_SHOW_URL = 0x8248
MAPI_ONLINE_PASSWORD = 0x8249
MAPI_APPOINTMENT_PROPOSED_DURATION = 0x8256
MAPI_APPT_COUNTER_PROPOSAL = 0x8257
MAPI_APPOINTMENT_PROPOSAL_NUMBER = 0x8259
MAPI_APPOINTMENT_NOT_ALLOW_PROPOSE = 0x825A
MAPI_APPT_TZDEF_START_DISPLAY = 0x825E
MAPI_APPT_TZDEF_END_DISPLAY = 0x825F
MAPI_APPT_TZDEF_RECUR = 0x8260
MAPI_REMINDER_MINUTES_BEFORE_START = 0x8501
MAPI_REMINDER_TIME = 0x8502
MAPI_REMINDER_SET = 0x8503
MAPI_PRIVATE = 0x8506
MAPI_AGING_DONT_AGE_ME = 0x850E
MAPI_FORM_STORAGE = 0x850F
MAPI_SIDE_EFFECTS = 0x8510
MAPI_REMOTE_STATUS = 0x8511
MAPI_PAGE_DIR_STREAM = 0x8513
MAPI_SMART_NO_ATTACH = 0x8514
MAPI_COMMON_START = 0x8516
MAPI_COMMON_END = 0x8517
MAPI_TASK_MODE = 0x8518
MAPI_FORM_PROP_STREAM = 0x851B
MAPI_REQUEST = 0x8530
MAPI_NON_SENDABLE_TO = 0x8536
MAPI_NON_SENDABLE_CC = 0x8537
MAPI_NON_SENDABLE_BCC = 0x8538
MAPI_COMPANIES = 0x8539
MAPI_CONTACTS = 0x853A
MAPI_PROP_DEF_STREAM = 0x8540
MAPI_SCRIPT_STREAM = 0x8541
MAPI_CUSTOM_FLAG = 0x8542
MAPI_OUTLOOK_CURRENT_VERSION = 0x8552
MAPI_CURRENT_VERSION_NAME = 0x8554
MAPI_REMINDER_NEXT_TIME = 0x8560
MAPI_HEADER_ITEM = 0x8578
MAPI_USE_TNEF = 0x8582
MAPI_TO_DO_TITLE = 0x85A4
MAPI_VALID_FLAG_STRING_PROOF = 0x85BF
MAPI_LOG_TYPE = 0x8700
MAPI_LOG_START = 0x8706
MAPI_LOG_DURATION = 0x8707
MAPI_LOG_END = 0x8708

CODE_TO_NAME = {
PidTagAccess: "PidTagAccess",
PidTagAccessControlListData: "PidTagAccessControlListData",
PidTagAccessLevel: "PidTagAccessLevel",
PidTagAccount: "PidTagAccount",
PidTagAdditionalRenEntryIds: "PidTagAdditionalRenEntryIds",
PidTagAdditionalRenEntryIdsEx: "PidTagAdditionalRenEntryIdsEx",
PidTagAddressBookAuthorizedSenders: "PidTagAddressBookAuthorizedSenders",
PidTagAddressBookContainerId: "PidTagAddressBookContainerId",
PidTagAddressBookDeliveryContentLength: "PidTagAddressBookDeliveryContentLength",
PidTagAddressBookDisplayNamePrintable: "PidTagAddressBookDisplayNamePrintable",
PidTagAddressBookDisplayTypeExtended: "PidTagAddressBookDisplayTypeExtended",
PidTagAddressBookDistributionListExternalMemberCount: "PidTagAddressBookDistributionListExternalMemberCount",
PidTagAddressBookDistributionListMemberCount: "PidTagAddressBookDistributionListMemberCount",
PidTagAddressBookDistributionListMemberSubmitAccepted: "PidTagAddressBookDistributionListMemberSubmitAccepted",
PidTagAddressBookDistributionListMemberSubmitRejected: "PidTagAddressBookDistributionListMemberSubmitRejected",
PidTagAddressBookDistributionListRejectMessagesFromDLMembers: "PidTagAddressBookDistributionListRejectMessagesFromDLMembers",
PidTagAddressBookEntryId: "PidTagAddressBookEntryId",
PidTagAddressBookExtensionAttribute1: "PidTagAddressBookExtensionAttribute1",
PidTagAddressBookExtensionAttribute10: "PidTagAddressBookExtensionAttribute10",
PidTagAddressBookExtensionAttribute11: "PidTagAddressBookExtensionAttribute11",
PidTagAddressBookExtensionAttribute12: "PidTagAddressBookExtensionAttribute12",
PidTagAddressBookExtensionAttribute13: "PidTagAddressBookExtensionAttribute13",
PidTagAddressBookExtensionAttribute14: "PidTagAddressBookExtensionAttribute14",
PidTagAddressBookExtensionAttribute15: "PidTagAddressBookExtensionAttribute15",
PidTagAddressBookExtensionAttribute2: "PidTagAddressBookExtensionAttribute2",
PidTagAddressBookExtensionAttribute3: "PidTagAddressBookExtensionAttribute3",
PidTagAddressBookExtensionAttribute4: "PidTagAddressBookExtensionAttribute4",
PidTagAddressBookExtensionAttribute5: "PidTagAddressBookExtensionAttribute5",
PidTagAddressBookExtensionAttribute6: "PidTagAddressBookExtensionAttribute6",
PidTagAddressBookExtensionAttribute7: "PidTagAddressBookExtensionAttribute7",
PidTagAddressBookExtensionAttribute8: "PidTagAddressBookExtensionAttribute8",
PidTagAddressBookExtensionAttribute9: "PidTagAddressBookExtensionAttribute9",
PidTagAddressBookFolderPathname: "PidTagAddressBookFolderPathname",
PidTagAddressBookHierarchicalChildDepartments: "PidTagAddressBookHierarchicalChildDepartments",
PidTagAddressBookHierarchicalDepartmentMembers: "PidTagAddressBookHierarchicalDepartmentMembers",
PidTagAddressBookHierarchicalIsHierarchicalGroup: "PidTagAddressBookHierarchicalIsHierarchicalGroup",
PidTagAddressBookHierarchicalParentDepartment: "PidTagAddressBookHierarchicalParentDepartment",
PidTagAddressBookHierarchicalRootDepartment: "PidTagAddressBookHierarchicalRootDepartment",
PidTagAddressBookHierarchicalShowInDepartments: "PidTagAddressBookHierarchicalShowInDepartments",
PidTagAddressBookHomeMessageDatabase: "PidTagAddressBookHomeMessageDatabase",
PidTagAddressBookIsMaster: "PidTagAddressBookIsMaster",
PidTagAddressBookIsMemberOfDistributionList: "PidTagAddressBookIsMemberOfDistributionList",
PidTagAddressBookManageDistributionList: "PidTagAddressBookManageDistributionList",
PidTagAddressBookManager: "PidTagAddressBookManager",
PidTagAddressBookManagerDistinguishedName: "PidTagAddressBookManagerDistinguishedName",
PidTagAddressBookMember: "PidTagAddressBookMember",
PidTagAddressBookMessageId: "PidTagAddressBookMessageId",
PidTagAddressBookModerationEnabled: "PidTagAddressBookModerationEnabled",
PidTagAddressBookNetworkAddress: "PidTagAddressBookNetworkAddress",
PidTagAddressBookObjectDistinguishedName: "PidTagAddressBookObjectDistinguishedName",
PidTagAddressBookObjectGuid: "PidTagAddressBookObjectGuid",
PidTagAddressBookOrganizationalUnitRootDistinguishedName: "PidTagAddressBookOrganizationalUnitRootDistinguishedName",
PidTagAddressBookOwner: "PidTagAddressBookOwner",
PidTagAddressBookOwnerBackLink: "PidTagAddressBookOwnerBackLink",
PidTagAddressBookParentEntryId: "PidTagAddressBookParentEntryId",
PidTagAddressBookPhoneticCompanyName: "PidTagAddressBookPhoneticCompanyName",
PidTagAddressBookPhoneticDepartmentName: "PidTagAddressBookPhoneticDepartmentName",
PidTagAddressBookPhoneticDisplayName: "PidTagAddressBookPhoneticDisplayName",
PidTagAddressBookPhoneticGivenName: "PidTagAddressBookPhoneticGivenName",
PidTagAddressBookPhoneticSurname: "PidTagAddressBookPhoneticSurname",
PidTagAddressBookProxyAddresses: "PidTagAddressBookProxyAddresses",
PidTagAddressBookPublicDelegates: "PidTagAddressBookPublicDelegates",
PidTagAddressBookReports: "PidTagAddressBookReports",
PidTagAddressBookRoomCapacity: "PidTagAddressBookRoomCapacity",
PidTagAddressBookRoomContainers: "PidTagAddressBookRoomContainers",
PidTagAddressBookRoomDescription: "PidTagAddressBookRoomDescription",
PidTagAddressBookSenderHintTranslations: "PidTagAddressBookSenderHintTranslations",
PidTagAddressBookSeniorityIndex: "PidTagAddressBookSeniorityIndex",
PidTagAddressBookTargetAddress: "PidTagAddressBookTargetAddress",
PidTagAddressBookUnauthorizedSenders: "PidTagAddressBookUnauthorizedSenders",
PidTagAddressBookX509Certificate: "PidTagAddressBookX509Certificate",
PidTagAddressType: "PidTagAddressType",
PidTagAlternateRecipientAllowed: "PidTagAlternateRecipientAllowed",
PidTagAnr: "PidTagAnr",
PidTagArchiveDate: "PidTagArchiveDate",
PidTagArchivePeriod: "PidTagArchivePeriod",
PidTagArchiveTag: "PidTagArchiveTag",
PidTagAssistant: "PidTagAssistant",
PidTagAssistantTelephoneNumber: "PidTagAssistantTelephoneNumber",
PidTagAssociated: "PidTagAssociated",
PidTagAttachAdditionalInformation: "PidTagAttachAdditionalInformation",
PidTagAttachContentBase: "PidTagAttachContentBase",
PidTagAttachContentId: "PidTagAttachContentId",
PidTagAttachContentLocation: "PidTagAttachContentLocation",
PidTagAttachDataBinary: "PidTagAttachDataBinary",
PidTagAttachDataObject: "PidTagAttachDataObject",
PidTagAttachEncoding: "PidTagAttachEncoding",
PidTagAttachExtension: "PidTagAttachExtension",
PidTagAttachFilename: "PidTagAttachFilename",
PidTagAttachFlags: "PidTagAttachFlags",
PidTagAttachLongFilename: "PidTagAttachLongFilename",
PidTagAttachLongPathname: "PidTagAttachLongPathname",
PidTagAttachmentContactPhoto: "PidTagAttachmentContactPhoto",
PidTagAttachmentFlags: "PidTagAttachmentFlags",
PidTagAttachmentHidden: "PidTagAttachmentHidden",
PidTagAttachmentLinkId: "PidTagAttachmentLinkId",
PidTagAttachMethod: "PidTagAttachMethod",
PidTagAttachMimeTag: "PidTagAttachMimeTag",
PidTagAttachNumber: "PidTagAttachNumber",
PidTagAttachPathname: "PidTagAttachPathname",
PidTagAttachPayloadClass: "PidTagAttachPayloadClass",
PidTagAttachPayloadProviderGuidString: "PidTagAttachPayloadProviderGuidString",
PidTagAttachRendering: "PidTagAttachRendering",
PidTagAttachSize: "PidTagAttachSize",
PidTagAttachTag: "PidTagAttachTag",
PidTagAttachTransportName: "PidTagAttachTransportName",
PidTagAttributeHidden: "PidTagAttributeHidden",
PidTagAttributeReadOnly: "PidTagAttributeReadOnly",
PidTagAutoForwardComment: "PidTagAutoForwardComment",
PidTagAutoForwarded: "PidTagAutoForwarded",
PidTagAutoResponseSuppress: "PidTagAutoResponseSuppress",
PidTagBirthday: "PidTagBirthday",
PidTagBlockStatus: "PidTagBlockStatus",
PidTagBody: "PidTagBody",
PidTagBodyContentId: "PidTagBodyContentId",
PidTagBodyContentLocation: "PidTagBodyContentLocation",
PidTagBodyHtml: "PidTagBodyHtml",
PidTagBusiness2TelephoneNumber: "PidTagBusiness2TelephoneNumber",
PidTagBusiness2TelephoneNumbers: "PidTagBusiness2TelephoneNumbers",
PidTagBusinessFaxNumber: "PidTagBusinessFaxNumber",
PidTagBusinessHomePage: "PidTagBusinessHomePage",
PidTagBusinessTelephoneNumber: "PidTagBusinessTelephoneNumber",
PidTagCallbackTelephoneNumber: "PidTagCallbackTelephoneNumber",
PidTagCallId: "PidTagCallId",
PidTagCarTelephoneNumber: "PidTagCarTelephoneNumber",
PidTagCdoRecurrenceid: "PidTagCdoRecurrenceid",
PidTagChangeKey: "PidTagChangeKey",
PidTagChangeNumber: "PidTagChangeNumber",
PidTagChildrensNames: "PidTagChildrensNames",
PidTagClientActions: "PidTagClientActions",
PidTagClientSubmitTime: "PidTagClientSubmitTime",
PidTagCodePageId: "PidTagCodePageId",
PidTagComment: "PidTagComment",
PidTagCompanyMainTelephoneNumber: "PidTagCompanyMainTelephoneNumber",
PidTagCompanyName: "PidTagCompanyName",
PidTagComputerNetworkName: "PidTagComputerNetworkName",
PidTagConflictEntryId: "PidTagConflictEntryId",
PidTagContainerClass: "PidTagContainerClass",
PidTagContainerContents: "PidTagContainerContents",
PidTagContainerFlags: "PidTagContainerFlags",
PidTagContainerHierarchy: "PidTagContainerHierarchy",
PidTagContentCount: "PidTagContentCount",
PidTagContentFilterSpamConfidenceLevel: "PidTagContentFilterSpamConfidenceLevel",
PidTagContentUnreadCount: "PidTagContentUnreadCount",
PidTagConversationId: "PidTagConversationId",
PidTagConversationIndex: "PidTagConversationIndex",
PidTagConversationIndexTracking: "PidTagConversationIndexTracking",
PidTagConversationTopic: "PidTagConversationTopic",
PidTagCountry: "PidTagCountry",
PidTagCreationTime: "PidTagCreationTime",
PidTagCreatorEntryId: "PidTagCreatorEntryId",
PidTagCreatorName: "PidTagCreatorName",
PidTagCustomerId: "PidTagCustomerId",
PidTagDamBackPatched: "PidTagDamBackPatched",
PidTagDamOriginalEntryId: "PidTagDamOriginalEntryId",
PidTagDefaultPostMessageClass: "PidTagDefaultPostMessageClass",
PidTagDeferredActionMessageOriginalEntryId: "PidTagDeferredActionMessageOriginalEntryId",
PidTagDeferredDeliveryTime: "PidTagDeferredDeliveryTime",
PidTagDeferredSendNumber: "PidTagDeferredSendNumber",
PidTagDeferredSendTime: "PidTagDeferredSendTime",
PidTagDeferredSendUnits: "PidTagDeferredSendUnits",
PidTagDelegatedByRule: "PidTagDelegatedByRule",
PidTagDelegateFlags: "PidTagDelegateFlags",
PidTagDeleteAfterSubmit: "PidTagDeleteAfterSubmit",
PidTagDeletedCountTotal: "PidTagDeletedCountTotal",
PidTagDeletedOn: "PidTagDeletedOn",
PidTagDeliverTime: "PidTagDeliverTime",
PidTagDepartmentName: "PidTagDepartmentName",
PidTagDepth: "PidTagDepth",
PidTagDisplayBcc: "PidTagDisplayBcc",
PidTagDisplayCc: "PidTagDisplayCc",
PidTagDisplayName: "PidTagDisplayName",
PidTagDisplayNamePrefix: "PidTagDisplayNamePrefix",
PidTagDisplayTo: "PidTagDisplayTo",
PidTagDisplayType: "PidTagDisplayType",
PidTagDisplayTypeEx: "PidTagDisplayTypeEx",
PidTagEmailAddress: "PidTagEmailAddress",
PidTagEndDate: "PidTagEndDate",
PidTagEntryId: "PidTagEntryId",
PidTagExceptionEndTime: "PidTagExceptionEndTime",
PidTagExceptionReplaceTime: "PidTagExceptionReplaceTime",
PidTagExceptionStartTime: "PidTagExceptionStartTime",
PidTagExchangeNTSecurityDescriptor: "PidTagExchangeNTSecurityDescriptor",
PidTagExpiryNumber: "PidTagExpiryNumber",
PidTagExpiryTime: "PidTagExpiryTime",
PidTagExpiryUnits: "PidTagExpiryUnits",
PidTagExtendedFolderFlags: "PidTagExtendedFolderFlags",
PidTagExtendedRuleMessageActions: "PidTagExtendedRuleMessageActions",
PidTagExtendedRuleMessageCondition: "PidTagExtendedRuleMessageCondition",
PidTagExtendedRuleSizeLimit: "PidTagExtendedRuleSizeLimit",
PidTagFaxNumberOfPages: "PidTagFaxNumberOfPages",
PidTagFlagCompleteTime: "PidTagFlagCompleteTime",
PidTagFlagStatus: "PidTagFlagStatus",
PidTagFlatUrlName: "PidTagFlatUrlName",
PidTagFolderAssociatedContents: "PidTagFolderAssociatedContents",
PidTagFolderId: "PidTagFolderId",
PidTagFolderType: "PidTagFolderType",
PidTagFollowupIcon: "PidTagFollowupIcon",
PidTagFreeBusyCountMonths: "PidTagFreeBusyCountMonths",
PidTagFreeBusyEntryIds: "PidTagFreeBusyEntryIds",
PidTagFreeBusyMessageEmailAddress: "PidTagFreeBusyMessageEmailAddress",
PidTagFreeBusyPublishEnd: "PidTagFreeBusyPublishEnd",
PidTagFreeBusyPublishStart: "PidTagFreeBusyPublishStart",
PidTagFreeBusyRangeTimestamp: "PidTagFreeBusyRangeTimestamp",
PidTagFtpSite: "PidTagFtpSite",
PidTagGatewayNeedsToRefresh: "PidTagGatewayNeedsToRefresh",
PidTagGender: "PidTagGender",
PidTagGeneration: "PidTagGeneration",
PidTagGivenName: "PidTagGivenName",
PidTagGovernmentIdNumber: "PidTagGovernmentIdNumber",
PidTagHasAttachments: "PidTagHasAttachments",
PidTagHasDeferredActionMessages: "PidTagHasDeferredActionMessages",
PidTagHasNamedProperties: "PidTagHasNamedProperties",
PidTagHasRules: "PidTagHasRules",
PidTagHierarchyChangeNumber: "PidTagHierarchyChangeNumber",
PidTagHobbies: "PidTagHobbies",
PidTagHome2TelephoneNumber: "PidTagHome2TelephoneNumber",
PidTagHome2TelephoneNumbers: "PidTagHome2TelephoneNumbers",
PidTagHomeAddressCity: "PidTagHomeAddressCity",
PidTagHomeAddressCountry: "PidTagHomeAddressCountry",
PidTagHomeAddressPostalCode: "PidTagHomeAddressPostalCode",
PidTagHomeAddressPostOfficeBox: "PidTagHomeAddressPostOfficeBox",
PidTagHomeAddressStateOrProvince: "PidTagHomeAddressStateOrProvince",
PidTagHomeAddressStreet: "PidTagHomeAddressStreet",
PidTagHomeFaxNumber: "PidTagHomeFaxNumber",
PidTagHomeTelephoneNumber: "PidTagHomeTelephoneNumber",
PidTagHtml: "PidTagHtml",
PidTagICalendarEndTime: "PidTagICalendarEndTime",
PidTagICalendarReminderNextTime: "PidTagICalendarReminderNextTime",
PidTagICalendarStartTime: "PidTagICalendarStartTime",
PidTagIconIndex: "PidTagIconIndex",
PidTagImportance: "PidTagImportance",
PidTagInConflict: "PidTagInConflict",
PidTagInitialDetailsPane: "PidTagInitialDetailsPane",
PidTagInitials: "PidTagInitials",
PidTagInReplyToId: "PidTagInReplyToId",
PidTagInstanceKey: "PidTagInstanceKey",
PidTagInstanceNum: "PidTagInstanceNum",
PidTagInstID: "PidTagInstID",
PidTagInternetCodepage: "PidTagInternetCodepage",
PidTagInternetMailOverrideFormat: "PidTagInternetMailOverrideFormat",
PidTagInternetMessageId: "PidTagInternetMessageId",
PidTagInternetReferences: "PidTagInternetReferences",
PidTagIpmAppointmentEntryId: "PidTagIpmAppointmentEntryId",
PidTagIpmContactEntryId: "PidTagIpmContactEntryId",
PidTagIpmDraftsEntryId: "PidTagIpmDraftsEntryId",
PidTagIpmJournalEntryId: "PidTagIpmJournalEntryId",
PidTagIpmNoteEntryId: "PidTagIpmNoteEntryId",
PidTagIpmTaskEntryId: "PidTagIpmTaskEntryId",
PidTagIsdnNumber: "PidTagIsdnNumber",
PidTagJunkAddRecipientsToSafeSendersList: "PidTagJunkAddRecipientsToSafeSendersList",
PidTagJunkIncludeContacts: "PidTagJunkIncludeContacts",
PidTagJunkPermanentlyDelete: "PidTagJunkPermanentlyDelete",
PidTagJunkPhishingEnableLinks: "PidTagJunkPhishingEnableLinks",
PidTagJunkThreshold: "PidTagJunkThreshold",
PidTagKeyword: "PidTagKeyword",
PidTagLanguage: "PidTagLanguage",
PidTagLastModificationTime: "PidTagLastModificationTime",
PidTagLastModifierEntryId: "PidTagLastModifierEntryId",
PidTagLastModifierName: "PidTagLastModifierName",
PidTagLastVerbExecuted: "PidTagLastVerbExecuted",
PidTagLastVerbExecutionTime: "PidTagLastVerbExecutionTime",
PidTagListHelp: "PidTagListHelp",
PidTagListSubscribe: "PidTagListSubscribe",
PidTagListUnsubscribe: "PidTagListUnsubscribe",
PidTagLocalCommitTime: "PidTagLocalCommitTime",
PidTagLocalCommitTimeMax: "PidTagLocalCommitTimeMax",
PidTagLocaleId: "PidTagLocaleId",
PidTagLocality: "PidTagLocality",
PidTagLocation: "PidTagLocation",
PidTagMailboxOwnerEntryId: "PidTagMailboxOwnerEntryId",
PidTagMailboxOwnerName: "PidTagMailboxOwnerName",
PidTagManagerName: "PidTagManagerName",
PidTagMappingSignature: "PidTagMappingSignature",
PidTagMaximumSubmitMessageSize: "PidTagMaximumSubmitMessageSize",
PidTagMemberId: "PidTagMemberId",
PidTagMemberName: "PidTagMemberName",
PidTagMemberRights: "PidTagMemberRights",
PidTagMessageAttachments: "PidTagMessageAttachments",
PidTagMessageCcMe: "PidTagMessageCcMe",
PidTagMessageClass: "PidTagMessageClass",
PidTagMessageCodepage: "PidTagMessageCodepage",
PidTagMessageDeliveryTime: "PidTagMessageDeliveryTime",
PidTagMessageEditorFormat: "PidTagMessageEditorFormat",
PidTagMessageFlags: "PidTagMessageFlags",
PidTagMessageHandlingSystemCommonName: "PidTagMessageHandlingSystemCommonName",
PidTagMessageLocaleId: "PidTagMessageLocaleId",
PidTagMessageRecipientMe: "PidTagMessageRecipientMe",
PidTagMessageRecipients: "PidTagMessageRecipients",
PidTagMessageSize: "PidTagMessageSize",
PidTagMessageSizeExtended: "PidTagMessageSizeExtended",
PidTagMessageStatus: "PidTagMessageStatus",
PidTagMessageSubmissionId: "PidTagMessageSubmissionId",
PidTagMessageToMe: "PidTagMessageToMe",
PidTagMid: "PidTagMid",
PidTagMiddleName: "PidTagMiddleName",
PidTagMimeSkeleton: "PidTagMimeSkeleton",
PidTagMobileTelephoneNumber: "PidTagMobileTelephoneNumber",
PidTagNativeBody: "PidTagNativeBody",
PidTagNextSendAcct: "PidTagNextSendAcct",
PidTagNickname: "PidTagNickname",
PidTagNonDeliveryReportDiagCode: "PidTagNonDeliveryReportDiagCode",
PidTagNonDeliveryReportReasonCode: "PidTagNonDeliveryReportReasonCode",
PidTagNonDeliveryReportStatusCode: "PidTagNonDeliveryReportStatusCode",
PidTagNonReceiptNotificationRequested: "PidTagNonReceiptNotificationRequested",
PidTagNormalizedSubject: "PidTagNormalizedSubject",
PidTagObjectType: "PidTagObjectType",
PidTagOfficeLocation: "PidTagOfficeLocation",
PidTagOfflineAddressBookContainerGuid: "PidTagOfflineAddressBookContainerGuid",
PidTagOfflineAddressBookDistinguishedName: "PidTagOfflineAddressBookDistinguishedName",
PidTagOfflineAddressBookMessageClass: "PidTagOfflineAddressBookMessageClass",
PidTagOfflineAddressBookName: "PidTagOfflineAddressBookName",
PidTagOfflineAddressBookSequence: "PidTagOfflineAddressBookSequence",
PidTagOfflineAddressBookTruncatedProperties: "PidTagOfflineAddressBookTruncatedProperties",
PidTagOrdinalMost: "PidTagOrdinalMost",
PidTagOrganizationalIdNumber: "PidTagOrganizationalIdNumber",
PidTagOriginalAuthorEntryId: "PidTagOriginalAuthorEntryId",
PidTagOriginalAuthorName: "PidTagOriginalAuthorName",
PidTagOriginalDeliveryTime: "PidTagOriginalDeliveryTime",
PidTagOriginalDisplayBcc: "PidTagOriginalDisplayBcc",
PidTagOriginalDisplayCc: "PidTagOriginalDisplayCc",
PidTagOriginalDisplayTo: "PidTagOriginalDisplayTo",
PidTagOriginalEntryId: "PidTagOriginalEntryId",
PidTagOriginalMessageClass: "PidTagOriginalMessageClass",
PidTagOriginalMessageId: "PidTagOriginalMessageId",
PidTagOriginalSenderAddressType: "PidTagOriginalSenderAddressType",
PidTagOriginalSenderEmailAddress: "PidTagOriginalSenderEmailAddress",
PidTagOriginalSenderEntryId: "PidTagOriginalSenderEntryId",
PidTagOriginalSenderName: "PidTagOriginalSenderName",
PidTagOriginalSenderSearchKey: "PidTagOriginalSenderSearchKey",
PidTagOriginalSensitivity: "PidTagOriginalSensitivity",
PidTagOriginalSentRepresentingAddressType: "PidTagOriginalSentRepresentingAddressType",
PidTagOriginalSentRepresentingEmailAddress: "PidTagOriginalSentRepresentingEmailAddress",
PidTagOriginalSentRepresentingEntryId: "PidTagOriginalSentRepresentingEntryId",
PidTagOriginalSentRepresentingName: "PidTagOriginalSentRepresentingName",
PidTagOriginalSentRepresentingSearchKey: "PidTagOriginalSentRepresentingSearchKey",
PidTagOriginalSubject: "PidTagOriginalSubject",
PidTagOriginalSubmitTime: "PidTagOriginalSubmitTime",
PidTagOriginatorDeliveryReportRequested: "PidTagOriginatorDeliveryReportRequested",
PidTagOriginatorNonDeliveryReportRequested: "PidTagOriginatorNonDeliveryReportRequested",
PidTagOscSyncEnabled: "PidTagOscSyncEnabled",
PidTagOtherAddressCity: "PidTagOtherAddressCity",
PidTagOtherAddressCountry: "PidTagOtherAddressCountry",
PidTagOtherAddressPostalCode: "PidTagOtherAddressPostalCode",
PidTagOtherAddressPostOfficeBox: "PidTagOtherAddressPostOfficeBox",
PidTagOtherAddressStateOrProvince: "PidTagOtherAddressStateOrProvince",
PidTagOtherAddressStreet: "PidTagOtherAddressStreet",
PidTagOtherTelephoneNumber: "PidTagOtherTelephoneNumber",
PidTagOutOfOfficeState: "PidTagOutOfOfficeState",
PidTagOwnerAppointmentId: "PidTagOwnerAppointmentId",
PidTagPagerTelephoneNumber: "PidTagPagerTelephoneNumber",
PidTagParentEntryId: "PidTagParentEntryId",
PidTagParentFolderId: "PidTagParentFolderId",
PidTagParentKey: "PidTagParentKey",
PidTagParentSourceKey: "PidTagParentSourceKey",
PidTagPersonalHomePage: "PidTagPersonalHomePage",
PidTagPolicyTag: "PidTagPolicyTag",
PidTagPostalAddress: "PidTagPostalAddress",
PidTagPostalCode: "PidTagPostalCode",
PidTagPostOfficeBox: "PidTagPostOfficeBox",
PidTagPredecessorChangeList: "PidTagPredecessorChangeList",
PidTagPrimaryFaxNumber: "PidTagPrimaryFaxNumber",
PidTagPrimarySendAccount: "PidTagPrimarySendAccount",
PidTagPrimaryTelephoneNumber: "PidTagPrimaryTelephoneNumber",
PidTagPriority: "PidTagPriority",
PidTagProcessed: "PidTagProcessed",
PidTagProfession: "PidTagProfession",
PidTagProhibitReceiveQuota: "PidTagProhibitReceiveQuota",
PidTagProhibitSendQuota: "PidTagProhibitSendQuota",
PidTagPurportedSenderDomain: "PidTagPurportedSenderDomain",
PidTagRadioTelephoneNumber: "PidTagRadioTelephoneNumber",
PidTagRead: "PidTagRead",
PidTagReadReceiptAddressType: "PidTagReadReceiptAddressType",
PidTagReadReceiptEmailAddress: "PidTagReadReceiptEmailAddress",
PidTagReadReceiptEntryId: "PidTagReadReceiptEntryId",
PidTagReadReceiptName: "PidTagReadReceiptName",
PidTagReadReceiptRequested: "PidTagReadReceiptRequested",
PidTagReadReceiptSearchKey: "PidTagReadReceiptSearchKey",
PidTagReadReceiptSmtpAddress: "PidTagReadReceiptSmtpAddress",
PidTagReceiptTime: "PidTagReceiptTime",
PidTagReceivedByAddressType: "PidTagReceivedByAddressType",
PidTagReceivedByEmailAddress: "PidTagReceivedByEmailAddress",
PidTagReceivedByEntryId: "PidTagReceivedByEntryId",
PidTagReceivedByName: "PidTagReceivedByName",
PidTagReceivedBySearchKey: "PidTagReceivedBySearchKey",
PidTagReceivedBySmtpAddress: "PidTagReceivedBySmtpAddress",
PidTagReceivedRepresentingAddressType: "PidTagReceivedRepresentingAddressType",
PidTagReceivedRepresentingEmailAddress: "PidTagReceivedRepresentingEmailAddress",
PidTagReceivedRepresentingEntryId: "PidTagReceivedRepresentingEntryId",
PidTagReceivedRepresentingName: "PidTagReceivedRepresentingName",
PidTagReceivedRepresentingSearchKey: "PidTagReceivedRepresentingSearchKey",
PidTagReceivedRepresentingSmtpAddress: "PidTagReceivedRepresentingSmtpAddress",
PidTagRecipientDisplayName: "PidTagRecipientDisplayName",
PidTagRecipientEntryId: "PidTagRecipientEntryId",
PidTagRecipientFlags: "PidTagRecipientFlags",
PidTagRecipientOrder: "PidTagRecipientOrder",
PidTagRecipientProposed: "PidTagRecipientProposed",
PidTagRecipientProposedEndTime: "PidTagRecipientProposedEndTime",
PidTagRecipientProposedStartTime: "PidTagRecipientProposedStartTime",
PidTagRecipientReassignmentProhibited: "PidTagRecipientReassignmentProhibited",
PidTagRecipientTrackStatus: "PidTagRecipientTrackStatus",
PidTagRecipientTrackStatusTime: "PidTagRecipientTrackStatusTime",
PidTagRecipientType: "PidTagRecipientType",
PidTagRecordKey: "PidTagRecordKey",
PidTagReferredByName: "PidTagReferredByName",
PidTagRemindersOnlineEntryId: "PidTagRemindersOnlineEntryId",
PidTagRemoteMessageTransferAgent: "PidTagRemoteMessageTransferAgent",
PidTagRenderingPosition: "PidTagRenderingPosition",
PidTagReplyRecipientEntries: "PidTagReplyRecipientEntries",
PidTagReplyRecipientNames: "PidTagReplyRecipientNames",
PidTagReplyRequested: "PidTagReplyRequested",
PidTagReplyTemplateId: "PidTagReplyTemplateId",
PidTagReplyTime: "PidTagReplyTime",
PidTagReportDisposition: "PidTagReportDisposition",
PidTagReportDispositionMode: "PidTagReportDispositionMode",
PidTagReportEntryId: "PidTagReportEntryId",
PidTagReportingMessageTransferAgent: "PidTagReportingMessageTransferAgent",
PidTagReportName: "PidTagReportName",
PidTagReportSearchKey: "PidTagReportSearchKey",
PidTagReportTag: "PidTagReportTag",
PidTagReportText: "PidTagReportText",
PidTagReportTime: "PidTagReportTime",
PidTagResolveMethod: "PidTagResolveMethod",
PidTagResponseRequested: "PidTagResponseRequested",
PidTagResponsibility: "PidTagResponsibility",
PidTagRetentionDate: "PidTagRetentionDate",
PidTagRetentionFlags: "PidTagRetentionFlags",
PidTagRetentionPeriod: "PidTagRetentionPeriod",
PidTagRights: "PidTagRights",
PidTagRoamingDatatypes: "PidTagRoamingDatatypes",
PidTagRoamingDictionary: "PidTagRoamingDictionary",
PidTagRoamingXmlStream: "PidTagRoamingXmlStream",
PidTagRowid: "PidTagRowid",
PidTagRowType: "PidTagRowType",
PidTagRtfCompressed: "PidTagRtfCompressed",
PidTagRtfInSync: "PidTagRtfInSync",
PidTagRuleActionNumber: "PidTagRuleActionNumber",
PidTagRuleActions: "PidTagRuleActions",
PidTagRuleActionType: "PidTagRuleActionType",
PidTagRuleCondition: "PidTagRuleCondition",
PidTagRuleError: "PidTagRuleError",
PidTagRuleFolderEntryId: "PidTagRuleFolderEntryId",
PidTagRuleId: "PidTagRuleId",
PidTagRuleIds: "PidTagRuleIds",
PidTagRuleLevel: "PidTagRuleLevel",
PidTagRuleMessageLevel: "PidTagRuleMessageLevel",
PidTagRuleMessageName: "PidTagRuleMessageName",
PidTagRuleMessageProvider: "PidTagRuleMessageProvider",
PidTagRuleMessageProviderData: "PidTagRuleMessageProviderData",
PidTagRuleMessageSequence: "PidTagRuleMessageSequence",
PidTagRuleMessageState: "PidTagRuleMessageState",
PidTagRuleMessageUserFlags: "PidTagRuleMessageUserFlags",
PidTagRuleName: "PidTagRuleName",
PidTagRuleProvider: "PidTagRuleProvider",
PidTagRuleProviderData: "PidTagRuleProviderData",
PidTagRuleSequence: "PidTagRuleSequence",
PidTagRuleState: "PidTagRuleState",
PidTagRuleUserFlags: "PidTagRuleUserFlags",
PidTagRwRulesStream: "PidTagRwRulesStream",
PidTagScheduleInfoAppointmentTombstone: "PidTagScheduleInfoAppointmentTombstone",
PidTagScheduleInfoAutoAcceptAppointments: "PidTagScheduleInfoAutoAcceptAppointments",
PidTagScheduleInfoDelegateEntryIds: "PidTagScheduleInfoDelegateEntryIds",
PidTagScheduleInfoDelegateNames: "PidTagScheduleInfoDelegateNames",
PidTagScheduleInfoDelegateNamesW: "PidTagScheduleInfoDelegateNamesW",
PidTagScheduleInfoDelegatorWantsCopy: "PidTagScheduleInfoDelegatorWantsCopy",
PidTagScheduleInfoDelegatorWantsInfo: "PidTagScheduleInfoDelegatorWantsInfo",
PidTagScheduleInfoDisallowOverlappingAppts: "PidTagScheduleInfoDisallowOverlappingAppts",
PidTagScheduleInfoDisallowRecurringAppts: "PidTagScheduleInfoDisallowRecurringAppts",
PidTagScheduleInfoDontMailDelegates: "PidTagScheduleInfoDontMailDelegates",
PidTagScheduleInfoFreeBusy: "PidTagScheduleInfoFreeBusy",
PidTagScheduleInfoFreeBusyAway: "PidTagScheduleInfoFreeBusyAway",
PidTagScheduleInfoFreeBusyBusy: "PidTagScheduleInfoFreeBusyBusy",
PidTagScheduleInfoFreeBusyMerged: "PidTagScheduleInfoFreeBusyMerged",
PidTagScheduleInfoFreeBusyTentative: "PidTagScheduleInfoFreeBusyTentative",
PidTagScheduleInfoMonthsAway: "PidTagScheduleInfoMonthsAway",
PidTagScheduleInfoMonthsBusy: "PidTagScheduleInfoMonthsBusy",
PidTagScheduleInfoMonthsMerged: "PidTagScheduleInfoMonthsMerged",
PidTagScheduleInfoMonthsTentative: "PidTagScheduleInfoMonthsTentative",
PidTagScheduleInfoResourceType: "PidTagScheduleInfoResourceType",
PidTagSchedulePlusFreeBusyEntryId: "PidTagSchedulePlusFreeBusyEntryId",
PidTagScriptData: "PidTagScriptData",
PidTagSearchFolderDefinition: "PidTagSearchFolderDefinition",
PidTagSearchFolderEfpFlags: "PidTagSearchFolderEfpFlags",
PidTagSearchFolderExpiration: "PidTagSearchFolderExpiration",
PidTagSearchFolderId: "PidTagSearchFolderId",
PidTagSearchFolderLastUsed: "PidTagSearchFolderLastUsed",
PidTagSearchFolderRecreateInfo: "PidTagSearchFolderRecreateInfo",
PidTagSearchFolderStorageType: "PidTagSearchFolderStorageType",
PidTagSearchFolderTag: "PidTagSearchFolderTag",
PidTagSearchFolderTemplateId: "PidTagSearchFolderTemplateId",
PidTagSearchKey: "PidTagSearchKey",
PidTagSecurityDescriptorAsXml: "PidTagSecurityDescriptorAsXml",
PidTagSelectable: "PidTagSelectable",
PidTagSenderAddressType: "PidTagSenderAddressType",
PidTagSenderEmailAddress: "PidTagSenderEmailAddress",
PidTagSenderEntryId: "PidTagSenderEntryId",
PidTagSenderIdStatus: "PidTagSenderIdStatus",
PidTagSenderName: "PidTagSenderName",
PidTagSenderSearchKey: "PidTagSenderSearchKey",
PidTagSenderSmtpAddress: "PidTagSenderSmtpAddress",
PidTagSenderTelephoneNumber: "PidTagSenderTelephoneNumber",
PidTagSendInternetEncoding: "PidTagSendInternetEncoding",
PidTagSendRichInfo: "PidTagSendRichInfo",
PidTagSensitivity: "PidTagSensitivity",
PidTagSentMailSvrEID: "PidTagSentMailSvrEID",
PidTagSentRepresentingAddressType: "PidTagSentRepresentingAddressType",
PidTagSentRepresentingEmailAddress: "PidTagSentRepresentingEmailAddress",
PidTagSentRepresentingEntryId: "PidTagSentRepresentingEntryId",
PidTagSentRepresentingFlags: "PidTagSentRepresentingFlags",
PidTagSentRepresentingName: "PidTagSentRepresentingName",
PidTagSentRepresentingSearchKey: "PidTagSentRepresentingSearchKey",
PidTagSentRepresentingSmtpAddress: "PidTagSentRepresentingSmtpAddress",
PidTagSmtpAddress: "PidTagSmtpAddress",
PidTagSortLocaleId: "PidTagSortLocaleId",
PidTagSourceKey: "PidTagSourceKey",
PidTagSpokenName: "PidTagSpokenName",
PidTagSpouseName: "PidTagSpouseName",
PidTagStartDate: "PidTagStartDate",
PidTagStartDateEtc: "PidTagStartDateEtc",
PidTagStateOrProvince: "PidTagStateOrProvince",
PidTagStoreEntryId: "PidTagStoreEntryId",
PidTagStoreState: "PidTagStoreState",
PidTagStoreSupportMask: "PidTagStoreSupportMask",
PidTagStreetAddress: "PidTagStreetAddress",
PidTagSubfolders: "PidTagSubfolders",
PidTagSubject: "PidTagSubject",
PidTagSubjectPrefix: "PidTagSubjectPrefix",
PidTagSupplementaryInfo: "PidTagSupplementaryInfo",
PidTagSurname: "PidTagSurname",
PidTagSwappedToDoData: "PidTagSwappedToDoData",
PidTagSwappedToDoStore: "PidTagSwappedToDoStore",
PidTagTargetEntryId: "PidTagTargetEntryId",
PidTagTelecommunicationsDeviceForDeafTelephoneNumber: "PidTagTelecommunicationsDeviceForDeafTelephoneNumber",
PidTagTelexNumber: "PidTagTelexNumber",
PidTagTemplateData: "PidTagTemplateData",
PidTagTemplateid: "PidTagTemplateid",
PidTagTextAttachmentCharset: "PidTagTextAttachmentCharset",
PidTagThumbnailPhoto: "PidTagThumbnailPhoto",
PidTagTitle: "PidTagTitle",
PidTagTnefCorrelationKey: "PidTagTnefCorrelationKey",
PidTagToDoItemFlags: "PidTagToDoItemFlags",
PidTagTransmittableDisplayName: "PidTagTransmittableDisplayName",
PidTagTransportMessageHeaders: "PidTagTransportMessageHeaders",
PidTagTrustSender: "PidTagTrustSender",
PidTagUserCertificate: "PidTagUserCertificate",
PidTagUserEntryId: "PidTagUserEntryId",
PidTagUserX509Certificate: "PidTagUserX509Certificate",
PidTagViewDescriptorBinary: "PidTagViewDescriptorBinary",
PidTagViewDescriptorName: "PidTagViewDescriptorName",
PidTagViewDescriptorStrings: "PidTagViewDescriptorStrings",
PidTagViewDescriptorVersion: "PidTagViewDescriptorVersion",
PidTagVoiceMessageAttachmentOrder: "PidTagVoiceMessageAttachmentOrder",
PidTagVoiceMessageDuration: "PidTagVoiceMessageDuration",
PidTagVoiceMessageSenderName: "PidTagVoiceMessageSenderName",
PidTagWeddingAnniversary: "PidTagWeddingAnniversary",
PidTagWlinkAddressBookEID: "PidTagWlinkAddressBookEID",
PidTagWlinkAddressBookStoreEID: "PidTagWlinkAddressBookStoreEID",
PidTagWlinkCalendarColor: "PidTagWlinkCalendarColor",
PidTagWlinkClientID: "PidTagWlinkClientID",
PidTagWlinkEntryId: "PidTagWlinkEntryId",
PidTagWlinkFlags: "PidTagWlinkFlags",
PidTagWlinkFolderType: "PidTagWlinkFolderType",
PidTagWlinkGroupClsid: "PidTagWlinkGroupClsid",
PidTagWlinkGroupHeaderID: "PidTagWlinkGroupHeaderID",
PidTagWlinkGroupName: "PidTagWlinkGroupName",
PidTagWlinkOrdinal: "PidTagWlinkOrdinal",
PidTagWlinkRecordKey: "PidTagWlinkRecordKey",
PidTagWlinkROGroupType: "PidTagWlinkROGroupType",
PidTagWlinkSaveStamp: "PidTagWlinkSaveStamp",
PidTagWlinkSection: "PidTagWlinkSection",
PidTagWlinkStoreEntryId: "PidTagWlinkStoreEntryId",
PidTagWlinkType: "PidTagWlinkType",

    MAPI_ACKNOWLEDGEMENT_MODE: "MAPI_ACKNOWLEDGEMENT_MODE",
    MAPI_ALTERNATE_RECIPIENT_ALLOWED: "MAPI_ALTERNATE_RECIPIENT_ALLOWED",
    MAPI_AUTHORIZING_USERS: "MAPI_AUTHORIZING_USERS",
    MAPI_AUTO_FORWARD_COMMENT: "MAPI_AUTO_FORWARD_COMMENT",
    MAPI_AUTO_FORWARDED: "MAPI_AUTO_FORWARDED",
    MAPI_CONTENT_CONFIDENTIALITY_ALGORITHM_ID: "MAPI_CONTENT_CONFIDENTIALITY_ALGORITHM_ID",
    MAPI_CONTENT_CORRELATOR: "MAPI_CONTENT_CORRELATOR",
    MAPI_CONTENT_IDENTIFIER: "MAPI_CONTENT_IDENTIFIER",
    MAPI_CONTENT_LENGTH: "MAPI_CONTENT_LENGTH",
    MAPI_CONTENT_RETURN_REQUESTED: "MAPI_CONTENT_RETURN_REQUESTED",
    MAPI_CONVERSATION_KEY: "MAPI_CONVERSATION_KEY",
    MAPI_CONVERSION_EITS: "MAPI_CONVERSION_EITS",
    MAPI_CONVERSION_WITH_LOSS_PROHIBITED: "MAPI_CONVERSION_WITH_LOSS_PROHIBITED",
    MAPI_CONVERTED_EITS: "MAPI_CONVERTED_EITS",
    MAPI_DEFERRED_DELIVERY_TIME: "MAPI_DEFERRED_DELIVERY_TIME",
    MAPI_DELIVER_TIME: "MAPI_DELIVER_TIME",
    MAPI_DISCARD_REASON: "MAPI_DISCARD_REASON",
    MAPI_DISCLOSURE_OF_RECIPIENTS: "MAPI_DISCLOSURE_OF_RECIPIENTS",
    MAPI_DL_EXPANSION_HISTORY: "MAPI_DL_EXPANSION_HISTORY",
    MAPI_DL_EXPANSION_PROHIBITED: "MAPI_DL_EXPANSION_PROHIBITED",
    MAPI_EXPIRY_TIME: "MAPI_EXPIRY_TIME",
    MAPI_IMPLICIT_CONVERSION_PROHIBITED: "MAPI_IMPLICIT_CONVERSION_PROHIBITED",
    MAPI_IMPORTANCE: "MAPI_IMPORTANCE",
    MAPI_IPM_ID: "MAPI_IPM_ID",
    MAPI_LATEST_DELIVERY_TIME: "MAPI_LATEST_DELIVERY_TIME",
    MAPI_MESSAGE_CLASS: "MAPI_MESSAGE_CLASS",
    MAPI_MESSAGE_DELIVERY_ID: "MAPI_MESSAGE_DELIVERY_ID",
    MAPI_MESSAGE_SECURITY_LABEL: "MAPI_MESSAGE_SECURITY_LABEL",
    MAPI_OBSOLETED_IPMS: "MAPI_OBSOLETED_IPMS",
    MAPI_ORIGINALLY_INTENDED_RECIPIENT_NAME: "MAPI_ORIGINALLY_INTENDED_RECIPIENT_NAME",
    MAPI_ORIGINAL_EITS: "MAPI_ORIGINAL_EITS",
    MAPI_ORIGINATOR_CERTIFICATE: "MAPI_ORIGINATOR_CERTIFICATE",
    MAPI_ORIGINATOR_DELIVERY_REPORT_REQUESTED: "MAPI_ORIGINATOR_DELIVERY_REPORT_REQUESTED",
    MAPI_ORIGINATOR_RETURN_ADDRESS: "MAPI_ORIGINATOR_RETURN_ADDRESS",
    MAPI_PARENT_KEY: "MAPI_PARENT_KEY",
    MAPI_PRIORITY: "MAPI_PRIORITY",
    MAPI_ORIGIN_CHECK: "MAPI_ORIGIN_CHECK",
    MAPI_PROOF_OF_SUBMISSION_REQUESTED: "MAPI_PROOF_OF_SUBMISSION_REQUESTED",
    MAPI_READ_RECEIPT_REQUESTED: "MAPI_READ_RECEIPT_REQUESTED",
    MAPI_RECEIPT_TIME: "MAPI_RECEIPT_TIME",
    MAPI_RECIPIENT_REASSIGNMENT_PROHIBITED: "MAPI_RECIPIENT_REASSIGNMENT_PROHIBITED",
    MAPI_REDIRECTION_HISTORY: "MAPI_REDIRECTION_HISTORY",
    MAPI_RELATED_IPMS: "MAPI_RELATED_IPMS",
    MAPI_ORIGINAL_SENSITIVITY: "MAPI_ORIGINAL_SENSITIVITY",
    MAPI_LANGUAGES: "MAPI_LANGUAGES",
    MAPI_REPLY_TIME: "MAPI_REPLY_TIME",
    MAPI_REPORT_TAG: "MAPI_REPORT_TAG",
    MAPI_REPORT_TIME: "MAPI_REPORT_TIME",
    MAPI_RETURNED_IPM: "MAPI_RETURNED_IPM",
    MAPI_SECURITY: "MAPI_SECURITY",
    MAPI_INCOMPLETE_COPY: "MAPI_INCOMPLETE_COPY",
    MAPI_SENSITIVITY: "MAPI_SENSITIVITY",
    MAPI_SUBJECT: "MAPI_SUBJECT",
    MAPI_SUBJECT_IPM: "MAPI_SUBJECT_IPM",
    MAPI_CLIENT_SUBMIT_TIME: "MAPI_CLIENT_SUBMIT_TIME",
    MAPI_REPORT_NAME: "MAPI_REPORT_NAME",
    MAPI_SENT_REPRESENTING_SEARCH_KEY: "MAPI_SENT_REPRESENTING_SEARCH_KEY",
    MAPI_X400_CONTENT_TYPE: "MAPI_X400_CONTENT_TYPE",
    MAPI_SUBJECT_PREFIX: "MAPI_SUBJECT_PREFIX",
    MAPI_NON_RECEIPT_REASON: "MAPI_NON_RECEIPT_REASON",
    MAPI_RECEIVED_BY_ENTRYID: "MAPI_RECEIVED_BY_ENTRYID",
    MAPI_RECEIVED_BY_NAME: "MAPI_RECEIVED_BY_NAME",
    MAPI_SENT_REPRESENTING_ENTRYID: "MAPI_SENT_REPRESENTING_ENTRYID",
    MAPI_SENT_REPRESENTING_NAME: "MAPI_SENT_REPRESENTING_NAME",
    MAPI_RCVD_REPRESENTING_ENTRYID: "MAPI_RCVD_REPRESENTING_ENTRYID",
    MAPI_RCVD_REPRESENTING_NAME: "MAPI_RCVD_REPRESENTING_NAME",
    MAPI_REPORT_ENTRYID: "MAPI_REPORT_ENTRYID",
    MAPI_READ_RECEIPT_ENTRYID: "MAPI_READ_RECEIPT_ENTRYID",
    MAPI_MESSAGE_SUBMISSION_ID: "MAPI_MESSAGE_SUBMISSION_ID",
    MAPI_PROVIDER_SUBMIT_TIME: "MAPI_PROVIDER_SUBMIT_TIME",
    MAPI_ORIGINAL_SUBJECT: "MAPI_ORIGINAL_SUBJECT",
    MAPI_DISC_VAL: "MAPI_DISC_VAL",
    MAPI_ORIG_MESSAGE_CLASS: "MAPI_ORIG_MESSAGE_CLASS",
    MAPI_ORIGINAL_AUTHOR_ENTRYID: "MAPI_ORIGINAL_AUTHOR_ENTRYID",
    MAPI_ORIGINAL_AUTHOR_NAME: "MAPI_ORIGINAL_AUTHOR_NAME",
    MAPI_ORIGINAL_SUBMIT_TIME: "MAPI_ORIGINAL_SUBMIT_TIME",
    MAPI_REPLY_RECIPIENT_ENTRIES: "MAPI_REPLY_RECIPIENT_ENTRIES",
    MAPI_REPLY_RECIPIENT_NAMES: "MAPI_REPLY_RECIPIENT_NAMES",
    MAPI_RECEIVED_BY_SEARCH_KEY: "MAPI_RECEIVED_BY_SEARCH_KEY",
    MAPI_RCVD_REPRESENTING_SEARCH_KEY: "MAPI_RCVD_REPRESENTING_SEARCH_KEY",
    MAPI_READ_RECEIPT_SEARCH_KEY: "MAPI_READ_RECEIPT_SEARCH_KEY",
    MAPI_REPORT_SEARCH_KEY: "MAPI_REPORT_SEARCH_KEY",
    MAPI_ORIGINAL_DELIVERY_TIME: "MAPI_ORIGINAL_DELIVERY_TIME",
    MAPI_ORIGINAL_AUTHOR_SEARCH_KEY: "MAPI_ORIGINAL_AUTHOR_SEARCH_KEY",
    MAPI_MESSAGE_TO_ME: "MAPI_MESSAGE_TO_ME",
    MAPI_MESSAGE_CC_ME: "MAPI_MESSAGE_CC_ME",
    MAPI_MESSAGE_RECIP_ME: "MAPI_MESSAGE_RECIP_ME",
    MAPI_ORIGINAL_SENDER_NAME: "MAPI_ORIGINAL_SENDER_NAME",
    MAPI_ORIGINAL_SENDER_ENTRYID: "MAPI_ORIGINAL_SENDER_ENTRYID",
    MAPI_ORIGINAL_SENDER_SEARCH_KEY: "MAPI_ORIGINAL_SENDER_SEARCH_KEY",
    MAPI_ORIGINAL_SENT_REPRESENTING_NAME: "MAPI_ORIGINAL_SENT_REPRESENTING_NAME",
    MAPI_ORIGINAL_SENT_REPRESENTING_ENTRYID: "MAPI_ORIGINAL_SENT_REPRESENTING_ENTRYID",
    MAPI_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY: "MAPI_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY",
    MAPI_START_DATE: "MAPI_START_DATE",
    MAPI_END_DATE: "MAPI_END_DATE",
    MAPI_OWNER_APPT_ID: "MAPI_OWNER_APPT_ID",
    MAPI_RESPONSE_REQUESTED: "MAPI_RESPONSE_REQUESTED",
    MAPI_SENT_REPRESENTING_ADDRTYPE: "MAPI_SENT_REPRESENTING_ADDRTYPE",
    MAPI_SENT_REPRESENTING_EMAIL_ADDRESS: "MAPI_SENT_REPRESENTING_EMAIL_ADDRESS",
    MAPI_ORIGINAL_SENDER_ADDRTYPE: "MAPI_ORIGINAL_SENDER_ADDRTYPE",
    MAPI_ORIGINAL_SENDER_EMAIL_ADDRESS: "MAPI_ORIGINAL_SENDER_EMAIL_ADDRESS",
    MAPI_ORIGINAL_SENT_REPRESENTING_ADDRTYPE: "MAPI_ORIGINAL_SENT_REPRESENTING_ADDRTYPE",
    MAPI_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS: "MAPI_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS",
    MAPI_CONVERSATION_TOPIC: "MAPI_CONVERSATION_TOPIC",
    MAPI_CONVERSATION_INDEX: "MAPI_CONVERSATION_INDEX",
    MAPI_ORIGINAL_DISPLAY_BCC: "MAPI_ORIGINAL_DISPLAY_BCC",
    MAPI_ORIGINAL_DISPLAY_CC: "MAPI_ORIGINAL_DISPLAY_CC",
    MAPI_ORIGINAL_DISPLAY_TO: "MAPI_ORIGINAL_DISPLAY_TO",
    MAPI_RECEIVED_BY_ADDRTYPE: "MAPI_RECEIVED_BY_ADDRTYPE",
    MAPI_RECEIVED_BY_EMAIL_ADDRESS: "MAPI_RECEIVED_BY_EMAIL_ADDRESS",
    MAPI_RCVD_REPRESENTING_ADDRTYPE: "MAPI_RCVD_REPRESENTING_ADDRTYPE",
    MAPI_RCVD_REPRESENTING_EMAIL_ADDRESS: "MAPI_RCVD_REPRESENTING_EMAIL_ADDRESS",
    MAPI_ORIGINAL_AUTHOR_ADDRTYPE: "MAPI_ORIGINAL_AUTHOR_ADDRTYPE",
    MAPI_ORIGINAL_AUTHOR_EMAIL_ADDRESS: "MAPI_ORIGINAL_AUTHOR_EMAIL_ADDRESS",
    MAPI_ORIGINALLY_INTENDED_RECIP_ADDRTYPE: "MAPI_ORIGINALLY_INTENDED_RECIP_ADDRTYPE",
    MAPI_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS: "MAPI_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS",
    MAPI_TRANSPORT_MESSAGE_HEADERS: "MAPI_TRANSPORT_MESSAGE_HEADERS",
    MAPI_DELEGATION: "MAPI_DELEGATION",
    MAPI_TNEF_CORRELATION_KEY: "MAPI_TNEF_CORRELATION_KEY",
    MAPI_CONTENT_INTEGRITY_CHECK: "MAPI_CONTENT_INTEGRITY_CHECK",
    MAPI_EXPLICIT_CONVERSION: "MAPI_EXPLICIT_CONVERSION",
    MAPI_IPM_RETURN_REQUESTED: "MAPI_IPM_RETURN_REQUESTED",
    MAPI_MESSAGE_TOKEN: "MAPI_MESSAGE_TOKEN",
    MAPI_NDR_REASON_CODE: "MAPI_NDR_REASON_CODE",
    MAPI_NDR_DIAG_CODE: "MAPI_NDR_DIAG_CODE",
    MAPI_NON_RECEIPT_NOTIFICATION_REQUESTED: "MAPI_NON_RECEIPT_NOTIFICATION_REQUESTED",
    MAPI_DELIVERY_POINT: "MAPI_DELIVERY_POINT",
    MAPI_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED: "MAPI_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED",
    MAPI_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT: "MAPI_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT",
    MAPI_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY: "MAPI_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY",
    MAPI_PHYSICAL_DELIVERY_MODE: "MAPI_PHYSICAL_DELIVERY_MODE",
    MAPI_PHYSICAL_DELIVERY_REPORT_REQUEST: "MAPI_PHYSICAL_DELIVERY_REPORT_REQUEST",
    MAPI_PHYSICAL_FORWARDING_ADDRESS: "MAPI_PHYSICAL_FORWARDING_ADDRESS",
    MAPI_PHYSICAL_FORWARDING_ADDRESS_REQUESTED: "MAPI_PHYSICAL_FORWARDING_ADDRESS_REQUESTED",
    MAPI_PHYSICAL_FORWARDING_PROHIBITED: "MAPI_PHYSICAL_FORWARDING_PROHIBITED",
    MAPI_PHYSICAL_RENDITION_ATTRIBUTES: "MAPI_PHYSICAL_RENDITION_ATTRIBUTES",
    MAPI_PROOF_OF_DELIVERY: "MAPI_PROOF_OF_DELIVERY",
    MAPI_PROOF_OF_DELIVERY_REQUESTED: "MAPI_PROOF_OF_DELIVERY_REQUESTED",
    MAPI_RECIPIENT_CERTIFICATE: "MAPI_RECIPIENT_CERTIFICATE",
    MAPI_RECIPIENT_NUMBER_FOR_ADVICE: "MAPI_RECIPIENT_NUMBER_FOR_ADVICE",
    MAPI_RECIPIENT_TYPE: "MAPI_RECIPIENT_TYPE",
    MAPI_REGISTERED_MAIL_TYPE: "MAPI_REGISTERED_MAIL_TYPE",
    MAPI_REPLY_REQUESTED: "MAPI_REPLY_REQUESTED",
    MAPI_REQUESTED_DELIVERY_METHOD: "MAPI_REQUESTED_DELIVERY_METHOD",
    MAPI_SENDER_ENTRYID: "MAPI_SENDER_ENTRYID",
    MAPI_SENDER_NAME: "MAPI_SENDER_NAME",
    MAPI_SUPPLEMENTARY_INFO: "MAPI_SUPPLEMENTARY_INFO",
    MAPI_TYPE_OF_MTS_USER: "MAPI_TYPE_OF_MTS_USER",
    MAPI_SENDER_SEARCH_KEY: "MAPI_SENDER_SEARCH_KEY",
    MAPI_SENDER_ADDRTYPE: "MAPI_SENDER_ADDRTYPE",
    MAPI_SENDER_EMAIL_ADDRESS: "MAPI_SENDER_EMAIL_ADDRESS",
    MAPI_CURRENT_VERSION: "MAPI_CURRENT_VERSION",
    MAPI_DELETE_AFTER_SUBMIT: "MAPI_DELETE_AFTER_SUBMIT",
    MAPI_DISPLAY_BCC: "MAPI_DISPLAY_BCC",
    MAPI_DISPLAY_CC: "MAPI_DISPLAY_CC",
    MAPI_DISPLAY_TO: "MAPI_DISPLAY_TO",
    MAPI_PARENT_DISPLAY: "MAPI_PARENT_DISPLAY",
    MAPI_MESSAGE_DELIVERY_TIME: "MAPI_MESSAGE_DELIVERY_TIME",
    MAPI_MESSAGE_FLAGS: "MAPI_MESSAGE_FLAGS",
    MAPI_MESSAGE_SIZE: "MAPI_MESSAGE_SIZE",
    MAPI_PARENT_ENTRYID: "MAPI_PARENT_ENTRYID",
    MAPI_SENTMAIL_ENTRYID: "MAPI_SENTMAIL_ENTRYID",
    MAPI_CORRELATE: "MAPI_CORRELATE",
    MAPI_CORRELATE_MTSID: "MAPI_CORRELATE_MTSID",
    MAPI_DISCRETE_VALUES: "MAPI_DISCRETE_VALUES",
    MAPI_RESPONSIBILITY: "MAPI_RESPONSIBILITY",
    MAPI_SPOOLER_STATUS: "MAPI_SPOOLER_STATUS",
    MAPI_TRANSPORT_STATUS: "MAPI_TRANSPORT_STATUS",
    MAPI_MESSAGE_RECIPIENTS: "MAPI_MESSAGE_RECIPIENTS",
    MAPI_MESSAGE_ATTACHMENTS: "MAPI_MESSAGE_ATTACHMENTS",
    MAPI_SUBMIT_FLAGS: "MAPI_SUBMIT_FLAGS",
    MAPI_RECIPIENT_STATUS: "MAPI_RECIPIENT_STATUS",
    MAPI_TRANSPORT_KEY: "MAPI_TRANSPORT_KEY",
    MAPI_MSG_STATUS: "MAPI_MSG_STATUS",
    MAPI_MESSAGE_DOWNLOAD_TIME: "MAPI_MESSAGE_DOWNLOAD_TIME",
    MAPI_CREATION_VERSION: "MAPI_CREATION_VERSION",
    MAPI_MODIFY_VERSION: "MAPI_MODIFY_VERSION",
    MAPI_HASATTACH: "MAPI_HASATTACH",
    MAPI_BODY_CRC: "MAPI_BODY_CRC",
    MAPI_NORMALIZED_SUBJECT: "MAPI_NORMALIZED_SUBJECT",
    MAPI_RTF_IN_SYNC: "MAPI_RTF_IN_SYNC",
    MAPI_ATTACH_SIZE: "MAPI_ATTACH_SIZE",
    MAPI_ATTACH_NUM: "MAPI_ATTACH_NUM",
    MAPI_PREPROCESS: "MAPI_PREPROCESS",
    MAPI_ORIGINATING_MTA_CERTIFICATE: "MAPI_ORIGINATING_MTA_CERTIFICATE",
    MAPI_PROOF_OF_SUBMISSION: "MAPI_PROOF_OF_SUBMISSION",
    MAPI_PRIMARY_SEND_ACCOUNT: "MAPI_PRIMARY_SEND_ACCOUNT",
    MAPI_NEXT_SEND_ACCT: "MAPI_NEXT_SEND_ACCT",
    MAPI_ACCESS: "MAPI_ACCESS",
    MAPI_ROW_TYPE: "MAPI_ROW_TYPE",
    MAPI_INSTANCE_KEY: "MAPI_INSTANCE_KEY",
    MAPI_ACCESS_LEVEL: "MAPI_ACCESS_LEVEL",
    MAPI_MAPPING_SIGNATURE: "MAPI_MAPPING_SIGNATURE",
    MAPI_RECORD_KEY: "MAPI_RECORD_KEY",
    MAPI_STORE_RECORD_KEY: "MAPI_STORE_RECORD_KEY",
    MAPI_STORE_ENTRYID: "MAPI_STORE_ENTRYID",
    MAPI_MINI_ICON: "MAPI_MINI_ICON",
    MAPI_ICON: "MAPI_ICON",
    MAPI_OBJECT_TYPE: "MAPI_OBJECT_TYPE",
    MAPI_ENTRYID: "MAPI_ENTRYID",
    MAPI_BODY: "MAPI_BODY",
    MAPI_REPORT_TEXT: "MAPI_REPORT_TEXT",
    MAPI_ORIGINATOR_AND_DL_EXPANSION_HISTORY: "MAPI_ORIGINATOR_AND_DL_EXPANSION_HISTORY",
    MAPI_REPORTING_DL_NAME: "MAPI_REPORTING_DL_NAME",
    MAPI_REPORTING_MTA_CERTIFICATE: "MAPI_REPORTING_MTA_CERTIFICATE",
    MAPI_RTF_SYNC_BODY_CRC: "MAPI_RTF_SYNC_BODY_CRC",
    MAPI_RTF_SYNC_BODY_COUNT: "MAPI_RTF_SYNC_BODY_COUNT",
    MAPI_RTF_SYNC_BODY_TAG: "MAPI_RTF_SYNC_BODY_TAG",
    MAPI_RTF_COMPRESSED: "MAPI_RTF_COMPRESSED",
    MAPI_RTF_SYNC_PREFIX_COUNT: "MAPI_RTF_SYNC_PREFIX_COUNT",
    MAPI_RTF_SYNC_TRAILING_COUNT: "MAPI_RTF_SYNC_TRAILING_COUNT",
    MAPI_ORIGINALLY_INTENDED_RECIP_ENTRYID: "MAPI_ORIGINALLY_INTENDED_RECIP_ENTRYID",
    MAPI_BODY_HTML: "MAPI_BODY_HTML",
    MAPI_NATIVE_BODY: "MAPI_NATIVE_BODY",
    MAPI_SMTP_MESSAGE_ID: "MAPI_SMTP_MESSAGE_ID",
    MAPI_INTERNET_REFERENCES: "MAPI_INTERNET_REFERENCES",
    MAPI_IN_REPLY_TO_ID: "MAPI_IN_REPLY_TO_ID",
    MAPI_INTERNET_RETURN_PATH: "MAPI_INTERNET_RETURN_PATH",
    MAPI_ICON_INDEX: "MAPI_ICON_INDEX",
    MAPI_LAST_VERB_EXECUTED: "MAPI_LAST_VERB_EXECUTED",
    MAPI_LAST_VERB_EXECUTION_TIME: "MAPI_LAST_VERB_EXECUTION_TIME",
    MAPI_URL_COMP_NAME: "MAPI_URL_COMP_NAME",
    MAPI_ATTRIBUTE_HIDDEN: "MAPI_ATTRIBUTE_HIDDEN",
    MAPI_ATTRIBUTE_SYSTEM: "MAPI_ATTRIBUTE_SYSTEM",
    MAPI_ATTRIBUTE_READ_ONLY: "MAPI_ATTRIBUTE_READ_ONLY",
    MAPI_ROWID: "MAPI_ROWID",
    MAPI_DISPLAY_NAME: "MAPI_DISPLAY_NAME",
    MAPI_ADDRTYPE: "MAPI_ADDRTYPE",
    MAPI_EMAIL_ADDRESS: "MAPI_EMAIL_ADDRESS",
    MAPI_COMMENT: "MAPI_COMMENT",
    MAPI_DEPTH: "MAPI_DEPTH",
    MAPI_PROVIDER_DISPLAY: "MAPI_PROVIDER_DISPLAY",
    MAPI_CREATION_TIME: "MAPI_CREATION_TIME",
    MAPI_LAST_MODIFICATION_TIME: "MAPI_LAST_MODIFICATION_TIME",
    MAPI_RESOURCE_FLAGS: "MAPI_RESOURCE_FLAGS",
    MAPI_PROVIDER_DLL_NAME: "MAPI_PROVIDER_DLL_NAME",
    MAPI_SEARCH_KEY: "MAPI_SEARCH_KEY",
    MAPI_PROVIDER_UID: "MAPI_PROVIDER_UID",
    MAPI_PROVIDER_ORDINAL: "MAPI_PROVIDER_ORDINAL",
    MAPI_TARGET_ENTRY_ID: "MAPI_TARGET_ENTRY_ID",
    MAPI_CONVERSATION_ID: "MAPI_CONVERSATION_ID",
    MAPI_CONVERSATION_INDEX_TRACKING: "MAPI_CONVERSATION_INDEX_TRACKING",
    MAPI_FORM_VERSION: "MAPI_FORM_VERSION",
    MAPI_FORM_CLSID: "MAPI_FORM_CLSID",
    MAPI_FORM_CONTACT_NAME: "MAPI_FORM_CONTACT_NAME",
    MAPI_FORM_CATEGORY: "MAPI_FORM_CATEGORY",
    MAPI_FORM_CATEGORY_SUB: "MAPI_FORM_CATEGORY_SUB",
    MAPI_FORM_HOST_MAP: "MAPI_FORM_HOST_MAP",
    MAPI_FORM_HIDDEN: "MAPI_FORM_HIDDEN",
    MAPI_FORM_DESIGNER_NAME: "MAPI_FORM_DESIGNER_NAME",
    MAPI_FORM_DESIGNER_GUID: "MAPI_FORM_DESIGNER_GUID",
    MAPI_FORM_MESSAGE_BEHAVIOR: "MAPI_FORM_MESSAGE_BEHAVIOR",
    MAPI_DEFAULT_STORE: "MAPI_DEFAULT_STORE",
    MAPI_STORE_SUPPORT_MASK: "MAPI_STORE_SUPPORT_MASK",
    MAPI_STORE_STATE: "MAPI_STORE_STATE",
    MAPI_STORE_UNICODE_MASK: "MAPI_STORE_UNICODE_MASK",
    MAPI_IPM_SUBTREE_SEARCH_KEY: "MAPI_IPM_SUBTREE_SEARCH_KEY",
    MAPI_IPM_OUTBOX_SEARCH_KEY: "MAPI_IPM_OUTBOX_SEARCH_KEY",
    MAPI_IPM_WASTEBASKET_SEARCH_KEY: "MAPI_IPM_WASTEBASKET_SEARCH_KEY",
    MAPI_IPM_SENTMAIL_SEARCH_KEY: "MAPI_IPM_SENTMAIL_SEARCH_KEY",
    MAPI_MDB_PROVIDER: "MAPI_MDB_PROVIDER",
    MAPI_RECEIVE_FOLDER_SETTINGS: "MAPI_RECEIVE_FOLDER_SETTINGS",
    MAPI_VALID_FOLDER_MASK: "MAPI_VALID_FOLDER_MASK",
    MAPI_IPM_SUBTREE_ENTRYID: "MAPI_IPM_SUBTREE_ENTRYID",
    MAPI_IPM_OUTBOX_ENTRYID: "MAPI_IPM_OUTBOX_ENTRYID",
    MAPI_IPM_WASTEBASKET_ENTRYID: "MAPI_IPM_WASTEBASKET_ENTRYID",
    MAPI_IPM_SENTMAIL_ENTRYID: "MAPI_IPM_SENTMAIL_ENTRYID",
    MAPI_VIEWS_ENTRYID: "MAPI_VIEWS_ENTRYID",
    MAPI_COMMON_VIEWS_ENTRYID: "MAPI_COMMON_VIEWS_ENTRYID",
    MAPI_FINDER_ENTRYID: "MAPI_FINDER_ENTRYID",
    MAPI_CONTAINER_FLAGS: "MAPI_CONTAINER_FLAGS",
    MAPI_FOLDER_TYPE: "MAPI_FOLDER_TYPE",
    MAPI_CONTENT_COUNT: "MAPI_CONTENT_COUNT",
    MAPI_CONTENT_UNREAD: "MAPI_CONTENT_UNREAD",
    MAPI_CREATE_TEMPLATES: "MAPI_CREATE_TEMPLATES",
    MAPI_DETAILS_TABLE: "MAPI_DETAILS_TABLE",
    MAPI_SEARCH: "MAPI_SEARCH",
    MAPI_SELECTABLE: "MAPI_SELECTABLE",
    MAPI_SUBFOLDERS: "MAPI_SUBFOLDERS",
    MAPI_STATUS: "MAPI_STATUS",
    MAPI_ANR: "MAPI_ANR",
    MAPI_CONTENTS_SORT_ORDER: "MAPI_CONTENTS_SORT_ORDER",
    MAPI_CONTAINER_HIERARCHY: "MAPI_CONTAINER_HIERARCHY",
    MAPI_CONTAINER_CONTENTS: "MAPI_CONTAINER_CONTENTS",
    MAPI_FOLDER_ASSOCIATED_CONTENTS: "MAPI_FOLDER_ASSOCIATED_CONTENTS",
    MAPI_DEF_CREATE_DL: "MAPI_DEF_CREATE_DL",
    MAPI_DEF_CREATE_MAILUSER: "MAPI_DEF_CREATE_MAILUSER",
    MAPI_CONTAINER_CLASS: "MAPI_CONTAINER_CLASS",
    MAPI_CONTAINER_MODIFY_VERSION: "MAPI_CONTAINER_MODIFY_VERSION",
    MAPI_AB_PROVIDER_ID: "MAPI_AB_PROVIDER_ID",
    MAPI_DEFAULT_VIEW_ENTRYID: "MAPI_DEFAULT_VIEW_ENTRYID",
    MAPI_ASSOC_CONTENT_COUNT: "MAPI_ASSOC_CONTENT_COUNT",
    MAPI_ATTACHMENT_X400_PARAMETERS: "MAPI_ATTACHMENT_X400_PARAMETERS",
    MAPI_ATTACH_DATA_OBJ: "MAPI_ATTACH_DATA_OBJ",
    MAPI_ATTACH_ENCODING: "MAPI_ATTACH_ENCODING",
    MAPI_ATTACH_EXTENSION: "MAPI_ATTACH_EXTENSION",
    MAPI_ATTACH_FILENAME: "MAPI_ATTACH_FILENAME",
    MAPI_ATTACH_METHOD: "MAPI_ATTACH_METHOD",
    MAPI_ATTACH_LONG_FILENAME: "MAPI_ATTACH_LONG_FILENAME",
    MAPI_ATTACH_PATHNAME: "MAPI_ATTACH_PATHNAME",
    MAPI_ATTACH_RENDERING: "MAPI_ATTACH_RENDERING",
    MAPI_ATTACH_TAG: "MAPI_ATTACH_TAG",
    MAPI_RENDERING_POSITION: "MAPI_RENDERING_POSITION",
    MAPI_ATTACH_TRANSPORT_NAME: "MAPI_ATTACH_TRANSPORT_NAME",
    MAPI_ATTACH_LONG_PATHNAME: "MAPI_ATTACH_LONG_PATHNAME",
    MAPI_ATTACH_MIME_TAG: "MAPI_ATTACH_MIME_TAG",
    MAPI_ATTACH_ADDITIONAL_INFO: "MAPI_ATTACH_ADDITIONAL_INFO",
    MAPI_ATTACH_MIME_SEQUENCE: "MAPI_ATTACH_MIME_SEQUENCE",
    MAPI_ATTACH_CONTENT_ID: "MAPI_ATTACH_CONTENT_ID",
    MAPI_ATTACH_CONTENT_LOCATION: "MAPI_ATTACH_CONTENT_LOCATION",
    MAPI_ATTACH_FLAGS: "MAPI_ATTACH_FLAGS",
    MAPI_DISPLAY_TYPE: "MAPI_DISPLAY_TYPE",
    MAPI_TEMPLATEID: "MAPI_TEMPLATEID",
    MAPI_PRIMARY_CAPABILITY: "MAPI_PRIMARY_CAPABILITY",
    MAPI_SMTP_ADDRESS: "MAPI_SMTP_ADDRESS",
    MAPI_7BIT_DISPLAY_NAME: "MAPI_7BIT_DISPLAY_NAME",
    MAPI_ACCOUNT: "MAPI_ACCOUNT",
    MAPI_ALTERNATE_RECIPIENT: "MAPI_ALTERNATE_RECIPIENT",
    MAPI_CALLBACK_TELEPHONE_NUMBER: "MAPI_CALLBACK_TELEPHONE_NUMBER",
    MAPI_CONVERSION_PROHIBITED: "MAPI_CONVERSION_PROHIBITED",
    MAPI_DISCLOSE_RECIPIENTS: "MAPI_DISCLOSE_RECIPIENTS",
    MAPI_GENERATION: "MAPI_GENERATION",
    MAPI_GIVEN_NAME: "MAPI_GIVEN_NAME",
    MAPI_GOVERNMENT_ID_NUMBER: "MAPI_GOVERNMENT_ID_NUMBER",
    MAPI_BUSINESS_TELEPHONE_NUMBER: "MAPI_BUSINESS_TELEPHONE_NUMBER",
    MAPI_HOME_TELEPHONE_NUMBER: "MAPI_HOME_TELEPHONE_NUMBER",
    MAPI_INITIALS: "MAPI_INITIALS",
    MAPI_KEYWORD: "MAPI_KEYWORD",
    MAPI_LANGUAGE: "MAPI_LANGUAGE",
    MAPI_LOCATION: "MAPI_LOCATION",
    MAPI_MAIL_PERMISSION: "MAPI_MAIL_PERMISSION",
    MAPI_MHS_COMMON_NAME: "MAPI_MHS_COMMON_NAME",
    MAPI_ORGANIZATIONAL_ID_NUMBER: "MAPI_ORGANIZATIONAL_ID_NUMBER",
    MAPI_SURNAME: "MAPI_SURNAME",
    MAPI_ORIGINAL_ENTRYID: "MAPI_ORIGINAL_ENTRYID",
    MAPI_ORIGINAL_DISPLAY_NAME: "MAPI_ORIGINAL_DISPLAY_NAME",
    MAPI_ORIGINAL_SEARCH_KEY: "MAPI_ORIGINAL_SEARCH_KEY",
    MAPI_POSTAL_ADDRESS: "MAPI_POSTAL_ADDRESS",
    MAPI_COMPANY_NAME: "MAPI_COMPANY_NAME",
    MAPI_TITLE: "MAPI_TITLE",
    MAPI_DEPARTMENT_NAME: "MAPI_DEPARTMENT_NAME",
    MAPI_OFFICE_LOCATION: "MAPI_OFFICE_LOCATION",
    MAPI_PRIMARY_TELEPHONE_NUMBER: "MAPI_PRIMARY_TELEPHONE_NUMBER",
    MAPI_BUSINESS2_TELEPHONE_NUMBER: "MAPI_BUSINESS2_TELEPHONE_NUMBER",
    MAPI_MOBILE_TELEPHONE_NUMBER: "MAPI_MOBILE_TELEPHONE_NUMBER",
    MAPI_RADIO_TELEPHONE_NUMBER: "MAPI_RADIO_TELEPHONE_NUMBER",
    MAPI_CAR_TELEPHONE_NUMBER: "MAPI_CAR_TELEPHONE_NUMBER",
    MAPI_OTHER_TELEPHONE_NUMBER: "MAPI_OTHER_TELEPHONE_NUMBER",
    MAPI_TRANSMITABLE_DISPLAY_NAME: "MAPI_TRANSMITABLE_DISPLAY_NAME",
    MAPI_PAGER_TELEPHONE_NUMBER: "MAPI_PAGER_TELEPHONE_NUMBER",
    MAPI_USER_CERTIFICATE: "MAPI_USER_CERTIFICATE",
    MAPI_PRIMARY_FAX_NUMBER: "MAPI_PRIMARY_FAX_NUMBER",
    MAPI_BUSINESS_FAX_NUMBER: "MAPI_BUSINESS_FAX_NUMBER",
    MAPI_HOME_FAX_NUMBER: "MAPI_HOME_FAX_NUMBER",
    MAPI_COUNTRY: "MAPI_COUNTRY",
    MAPI_LOCALITY: "MAPI_LOCALITY",
    MAPI_STATE_OR_PROVINCE: "MAPI_STATE_OR_PROVINCE",
    MAPI_STREET_ADDRESS: "MAPI_STREET_ADDRESS",
    MAPI_POSTAL_CODE: "MAPI_POSTAL_CODE",
    MAPI_POST_OFFICE_BOX: "MAPI_POST_OFFICE_BOX",
    MAPI_TELEX_NUMBER: "MAPI_TELEX_NUMBER",
    MAPI_ISDN_NUMBER: "MAPI_ISDN_NUMBER",
    MAPI_ASSISTANT_TELEPHONE_NUMBER: "MAPI_ASSISTANT_TELEPHONE_NUMBER",
    MAPI_HOME2_TELEPHONE_NUMBER: "MAPI_HOME2_TELEPHONE_NUMBER",
    MAPI_ASSISTANT: "MAPI_ASSISTANT",
    MAPI_SEND_RICH_INFO: "MAPI_SEND_RICH_INFO",
    MAPI_WEDDING_ANNIVERSARY: "MAPI_WEDDING_ANNIVERSARY",
    MAPI_BIRTHDAY: "MAPI_BIRTHDAY",
    MAPI_HOBBIES: "MAPI_HOBBIES",
    MAPI_MIDDLE_NAME: "MAPI_MIDDLE_NAME",
    MAPI_DISPLAY_NAME_PREFIX: "MAPI_DISPLAY_NAME_PREFIX",
    MAPI_PROFESSION: "MAPI_PROFESSION",
    MAPI_PREFERRED_BY_NAME: "MAPI_PREFERRED_BY_NAME",
    MAPI_SPOUSE_NAME: "MAPI_SPOUSE_NAME",
    MAPI_COMPUTER_NETWORK_NAME: "MAPI_COMPUTER_NETWORK_NAME",
    MAPI_CUSTOMER_ID: "MAPI_CUSTOMER_ID",
    MAPI_TTYTDD_PHONE_NUMBER: "MAPI_TTYTDD_PHONE_NUMBER",
    MAPI_FTP_SITE: "MAPI_FTP_SITE",
    MAPI_GENDER: "MAPI_GENDER",
    MAPI_MANAGER_NAME: "MAPI_MANAGER_NAME",
    MAPI_NICKNAME: "MAPI_NICKNAME",
    MAPI_PERSONAL_HOME_PAGE: "MAPI_PERSONAL_HOME_PAGE",
    MAPI_BUSINESS_HOME_PAGE: "MAPI_BUSINESS_HOME_PAGE",
    MAPI_CONTACT_VERSION: "MAPI_CONTACT_VERSION",
    MAPI_CONTACT_ENTRYIDS: "MAPI_CONTACT_ENTRYIDS",
    MAPI_CONTACT_ADDRTYPES: "MAPI_CONTACT_ADDRTYPES",
    MAPI_CONTACT_DEFAULT_ADDRESS_INDEX: "MAPI_CONTACT_DEFAULT_ADDRESS_INDEX",
    MAPI_CONTACT_EMAIL_ADDRESSES: "MAPI_CONTACT_EMAIL_ADDRESSES",
    MAPI_COMPANY_MAIN_PHONE_NUMBER: "MAPI_COMPANY_MAIN_PHONE_NUMBER",
    MAPI_CHILDRENS_NAMES: "MAPI_CHILDRENS_NAMES",
    MAPI_HOME_ADDRESS_CITY: "MAPI_HOME_ADDRESS_CITY",
    MAPI_HOME_ADDRESS_COUNTRY: "MAPI_HOME_ADDRESS_COUNTRY",
    MAPI_HOME_ADDRESS_POSTAL_CODE: "MAPI_HOME_ADDRESS_POSTAL_CODE",
    MAPI_HOME_ADDRESS_STATE_OR_PROVINCE: "MAPI_HOME_ADDRESS_STATE_OR_PROVINCE",
    MAPI_HOME_ADDRESS_STREET: "MAPI_HOME_ADDRESS_STREET",
    MAPI_HOME_ADDRESS_POST_OFFICE_BOX: "MAPI_HOME_ADDRESS_POST_OFFICE_BOX",
    MAPI_OTHER_ADDRESS_CITY: "MAPI_OTHER_ADDRESS_CITY",
    MAPI_OTHER_ADDRESS_COUNTRY: "MAPI_OTHER_ADDRESS_COUNTRY",
    MAPI_OTHER_ADDRESS_POSTAL_CODE: "MAPI_OTHER_ADDRESS_POSTAL_CODE",
    MAPI_OTHER_ADDRESS_STATE_OR_PROVINCE: "MAPI_OTHER_ADDRESS_STATE_OR_PROVINCE",
    MAPI_OTHER_ADDRESS_STREET: "MAPI_OTHER_ADDRESS_STREET",
    MAPI_OTHER_ADDRESS_POST_OFFICE_BOX: "MAPI_OTHER_ADDRESS_POST_OFFICE_BOX",
    MAPI_SEND_INTERNET_ENCODING: "MAPI_SEND_INTERNET_ENCODING",
    MAPI_STORE_PROVIDERS: "MAPI_STORE_PROVIDERS",
    MAPI_AB_PROVIDERS: "MAPI_AB_PROVIDERS",
    MAPI_TRANSPORT_PROVIDERS: "MAPI_TRANSPORT_PROVIDERS",
    MAPI_DEFAULT_PROFILE: "MAPI_DEFAULT_PROFILE",
    MAPI_AB_SEARCH_PATH: "MAPI_AB_SEARCH_PATH",
    MAPI_AB_DEFAULT_DIR: "MAPI_AB_DEFAULT_DIR",
    MAPI_AB_DEFAULT_PAB: "MAPI_AB_DEFAULT_PAB",
    MAPI_FILTERING_HOOKS: "MAPI_FILTERING_HOOKS",
    MAPI_SERVICE_NAME: "MAPI_SERVICE_NAME",
    MAPI_SERVICE_DLL_NAME: "MAPI_SERVICE_DLL_NAME",
    MAPI_SERVICE_ENTRY_NAME: "MAPI_SERVICE_ENTRY_NAME",
    MAPI_SERVICE_UID: "MAPI_SERVICE_UID",
    MAPI_SERVICE_EXTRA_UIDS: "MAPI_SERVICE_EXTRA_UIDS",
    MAPI_SERVICES: "MAPI_SERVICES",
    MAPI_SERVICE_SUPPORT_FILES: "MAPI_SERVICE_SUPPORT_FILES",
    MAPI_SERVICE_DELETE_FILES: "MAPI_SERVICE_DELETE_FILES",
    MAPI_AB_SEARCH_PATH_UPDATE: "MAPI_AB_SEARCH_PATH_UPDATE",
    MAPI_PROFILE_NAME: "MAPI_PROFILE_NAME",
    MAPI_IDENTITY_DISPLAY: "MAPI_IDENTITY_DISPLAY",
    MAPI_IDENTITY_ENTRYID: "MAPI_IDENTITY_ENTRYID",
    MAPI_RESOURCE_METHODS: "MAPI_RESOURCE_METHODS",
    MAPI_RESOURCE_TYPE: "MAPI_RESOURCE_TYPE",
    MAPI_STATUS_CODE: "MAPI_STATUS_CODE",
    MAPI_IDENTITY_SEARCH_KEY: "MAPI_IDENTITY_SEARCH_KEY",
    MAPI_OWN_STORE_ENTRYID: "MAPI_OWN_STORE_ENTRYID",
    MAPI_RESOURCE_PATH: "MAPI_RESOURCE_PATH",
    MAPI_STATUS_STRING: "MAPI_STATUS_STRING",
    MAPI_X400_DEFERRED_DELIVERY_CANCEL: "MAPI_X400_DEFERRED_DELIVERY_CANCEL",
    MAPI_HEADER_FOLDER_ENTRYID: "MAPI_HEADER_FOLDER_ENTRYID",
    MAPI_REMOTE_PROGRESS: "MAPI_REMOTE_PROGRESS",
    MAPI_REMOTE_PROGRESS_TEXT: "MAPI_REMOTE_PROGRESS_TEXT",
    MAPI_REMOTE_VALIDATE_OK: "MAPI_REMOTE_VALIDATE_OK",
    MAPI_CONTROL_FLAGS: "MAPI_CONTROL_FLAGS",
    MAPI_CONTROL_STRUCTURE: "MAPI_CONTROL_STRUCTURE",
    MAPI_CONTROL_TYPE: "MAPI_CONTROL_TYPE",
    MAPI_DELTAX: "MAPI_DELTAX",
    MAPI_DELTAY: "MAPI_DELTAY",
    MAPI_XPOS: "MAPI_XPOS",
    MAPI_YPOS: "MAPI_YPOS",
    MAPI_CONTROL_ID: "MAPI_CONTROL_ID",
    MAPI_INITIAL_DETAILS_PANE: "MAPI_INITIAL_DETAILS_PANE",
    MAPI_UNCOMPRESSED_BODY: "MAPI_UNCOMPRESSED_BODY",
    MAPI_INTERNET_CODEPAGE: "MAPI_INTERNET_CODEPAGE",
    MAPI_AUTO_RESPONSE_SUPPRESS: "MAPI_AUTO_RESPONSE_SUPPRESS",
    MAPI_MESSAGE_LOCALE_ID: "MAPI_MESSAGE_LOCALE_ID",
    MAPI_RULE_TRIGGER_HISTORY: "MAPI_RULE_TRIGGER_HISTORY",
    MAPI_MOVE_TO_STORE_ENTRYID: "MAPI_MOVE_TO_STORE_ENTRYID",
    MAPI_MOVE_TO_FOLDER_ENTRYID: "MAPI_MOVE_TO_FOLDER_ENTRYID",
    MAPI_STORAGE_QUOTA_LIMIT: "MAPI_STORAGE_QUOTA_LIMIT",
    MAPI_EXCESS_STORAGE_USED: "MAPI_EXCESS_STORAGE_USED",
    MAPI_SVR_GENERATING_QUOTA_MSG: "MAPI_SVR_GENERATING_QUOTA_MSG",
    MAPI_CREATOR_NAME: "MAPI_CREATOR_NAME",
    MAPI_CREATOR_ENTRY_ID: "MAPI_CREATOR_ENTRY_ID",
    MAPI_LAST_MODIFIER_NAME: "MAPI_LAST_MODIFIER_NAME",
    MAPI_LAST_MODIFIER_ENTRY_ID: "MAPI_LAST_MODIFIER_ENTRY_ID",
    MAPI_REPLY_RECIPIENT_SMTP_PROXIES: "MAPI_REPLY_RECIPIENT_SMTP_PROXIES",
    MAPI_MESSAGE_CODEPAGE: "MAPI_MESSAGE_CODEPAGE",
    MAPI_EXTENDED_ACL_DATA: "MAPI_EXTENDED_ACL_DATA",
    MAPI_SENDER_FLAGS: "MAPI_SENDER_FLAGS",
    MAPI_SENT_REPRESENTING_FLAGS: "MAPI_SENT_REPRESENTING_FLAGS",
    MAPI_RECEIVED_BY_FLAGS: "MAPI_RECEIVED_BY_FLAGS",
    MAPI_RECEIVED_REPRESENTING_FLAGS: "MAPI_RECEIVED_REPRESENTING_FLAGS",
    MAPI_CREATOR_ADDRESS_TYPE: "MAPI_CREATOR_ADDRESS_TYPE",
    MAPI_CREATOR_EMAIL_ADDRESS: "MAPI_CREATOR_EMAIL_ADDRESS",
    MAPI_SENDER_SIMPLE_DISPLAY_NAME: "MAPI_SENDER_SIMPLE_DISPLAY_NAME",
    MAPI_SENT_REPRESENTING_SIMPLE_DISPLAY_NAME: "MAPI_SENT_REPRESENTING_SIMPLE_DISPLAY_NAME",
    MAPI_RECEIVED_REPRESENTING_SIMPLE_DISPLAY_NAME: "MAPI_RECEIVED_REPRESENTING_SIMPLE_DISPLAY_NAME",
    MAPI_CREATOR_SIMPLE_DISP_NAME: "MAPI_CREATOR_SIMPLE_DISP_NAME",
    MAPI_LAST_MODIFIER_SIMPLE_DISPLAY_NAME: "MAPI_LAST_MODIFIER_SIMPLE_DISPLAY_NAME",
    MAPI_CONTENT_FILTER_SPAM_CONFIDENCE_LEVEL: "MAPI_CONTENT_FILTER_SPAM_CONFIDENCE_LEVEL",
    MAPI_INTERNET_MAIL_OVERRIDE_FORMAT: "MAPI_INTERNET_MAIL_OVERRIDE_FORMAT",
    MAPI_MESSAGE_EDITOR_FORMAT: "MAPI_MESSAGE_EDITOR_FORMAT",
    MAPI_SENDER_SMTP_ADDRESS: "MAPI_SENDER_SMTP_ADDRESS",
    MAPI_SENT_REPRESENTING_SMTP_ADDRESS: "MAPI_SENT_REPRESENTING_SMTP_ADDRESS",
    MAPI_READ_RECEIPT_SMTP_ADDRESS: "MAPI_READ_RECEIPT_SMTP_ADDRESS",
    MAPI_RECEIVED_BY_SMTP_ADDRESS: "MAPI_RECEIVED_BY_SMTP_ADDRESS",
    MAPI_RECEIVED_REPRESENTING_SMTP_ADDRESS: "MAPI_RECEIVED_REPRESENTING_SMTP_ADDRESS",
    MAPI_SENDING_SMTP_ADDRESS: "MAPI_SENDING_SMTP_ADDRESS",
    MAPI_SIP_ADDRESS: "MAPI_SIP_ADDRESS",
    MAPI_RECIPIENT_DISPLAY_NAME: "MAPI_RECIPIENT_DISPLAY_NAME",
    MAPI_RECIPIENT_ENTRYID: "MAPI_RECIPIENT_ENTRYID",
    MAPI_RECIPIENT_FLAGS: "MAPI_RECIPIENT_FLAGS",
    MAPI_RECIPIENT_TRACKSTATUS: "MAPI_RECIPIENT_TRACKSTATUS",
    MAPI_CHANGE_KEY: "MAPI_CHANGE_KEY",
    MAPI_PREDECESSOR_CHANGE_LIST: "MAPI_PREDECESSOR_CHANGE_LIST",
    MAPI_ID_SECURE_MIN: "MAPI_ID_SECURE_MIN",
    MAPI_ID_SECURE_MAX: "MAPI_ID_SECURE_MAX",
    MAPI_VOICE_MESSAGE_DURATION: "MAPI_VOICE_MESSAGE_DURATION",
    MAPI_SENDER_TELEPHONE_NUMBER: "MAPI_SENDER_TELEPHONE_NUMBER",
    MAPI_VOICE_MESSAGE_SENDER_NAME: "MAPI_VOICE_MESSAGE_SENDER_NAME",
    MAPI_FAX_NUMBER_OF_PAGES: "MAPI_FAX_NUMBER_OF_PAGES",
    MAPI_VOICE_MESSAGE_ATTACHMENT_ORDER: "MAPI_VOICE_MESSAGE_ATTACHMENT_ORDER",
    MAPI_CALL_ID: "MAPI_CALL_ID",
    MAPI_ATTACHMENT_LINK_ID: "MAPI_ATTACHMENT_LINK_ID",
    MAPI_EXCEPTION_START_TIME: "MAPI_EXCEPTION_START_TIME",
    MAPI_EXCEPTION_END_TIME: "MAPI_EXCEPTION_END_TIME",
    MAPI_ATTACHMENT_FLAGS: "MAPI_ATTACHMENT_FLAGS",
    MAPI_ATTACHMENT_HIDDEN: "MAPI_ATTACHMENT_HIDDEN",
    MAPI_ATTACHMENT_CONTACT_PHOTO: "MAPI_ATTACHMENT_CONTACT_PHOTO",
    MAPI_FILE_UNDER: "MAPI_FILE_UNDER",
    MAPI_FILE_UNDER_ID: "MAPI_FILE_UNDER_ID",
    MAPI_CONTACT_ITEM_DATA: "MAPI_CONTACT_ITEM_DATA",
    MAPI_REFERRED_BY: "MAPI_REFERRED_BY",
    MAPI_DEPARTMENT: "MAPI_DEPARTMENT",
    MAPI_HAS_PICTURE: "MAPI_HAS_PICTURE",
    MAPI_HOME_ADDRESS: "MAPI_HOME_ADDRESS",
    MAPI_WORK_ADDRESS: "MAPI_WORK_ADDRESS",
    MAPI_OTHER_ADDRESS: "MAPI_OTHER_ADDRESS",
    MAPI_POSTAL_ADDRESS_ID: "MAPI_POSTAL_ADDRESS_ID",
    MAPI_CONTACT_CHARACTER_SET: "MAPI_CONTACT_CHARACTER_SET",
    MAPI_AUTO_LOG: "MAPI_AUTO_LOG",
    MAPI_FILE_UNDER_LIST: "MAPI_FILE_UNDER_LIST",
    MAPI_EMAIL_LIST: "MAPI_EMAIL_LIST",
    MAPI_ADDRESS_BOOK_PROVIDER_EMAIL_LIST: "MAPI_ADDRESS_BOOK_PROVIDER_EMAIL_LIST",
    MAPI_ADDRESS_BOOK_PROVIDER_ARRAY_TYPE: "MAPI_ADDRESS_BOOK_PROVIDER_ARRAY_TYPE",
    MAPI_HTML: "MAPI_HTML",
    MAPI_YOMI_FIRST_NAME: "MAPI_YOMI_FIRST_NAME",
    MAPI_YOMI_LAST_NAME: "MAPI_YOMI_LAST_NAME",
    MAPI_YOMI_COMPANY_NAME: "MAPI_YOMI_COMPANY_NAME",
    MAPI_BUSINESS_CARD_DISPLAY_DEFINITION: "MAPI_BUSINESS_CARD_DISPLAY_DEFINITION",
    MAPI_BUSINESS_CARD_CARD_PICTURE: "MAPI_BUSINESS_CARD_CARD_PICTURE",
    MAPI_WORK_ADDRESS_STREET: "MAPI_WORK_ADDRESS_STREET",
    MAPI_WORK_ADDRESS_CITY: "MAPI_WORK_ADDRESS_CITY",
    MAPI_WORK_ADDRESS_STATE: "MAPI_WORK_ADDRESS_STATE",
    MAPI_WORK_ADDRESS_POSTAL_CODE: "MAPI_WORK_ADDRESS_POSTAL_CODE",
    MAPI_WORK_ADDRESS_COUNTRY: "MAPI_WORK_ADDRESS_COUNTRY",
    MAPI_WORK_ADDRESS_POST_OFFICE_BOX: "MAPI_WORK_ADDRESS_POST_OFFICE_BOX",
    MAPI_DISTRIBUTION_LIST_CHECKSUM: "MAPI_DISTRIBUTION_LIST_CHECKSUM",
    MAPI_BIRTHDAY_EVENT_ENTRY_ID: "MAPI_BIRTHDAY_EVENT_ENTRY_ID",
    MAPI_ANNIVERSARY_EVENT_ENTRY_ID: "MAPI_ANNIVERSARY_EVENT_ENTRY_ID",
    MAPI_CONTACT_USER_FIELD1: "MAPI_CONTACT_USER_FIELD1",
    MAPI_CONTACT_USER_FIELD2: "MAPI_CONTACT_USER_FIELD2",
    MAPI_CONTACT_USER_FIELD3: "MAPI_CONTACT_USER_FIELD3",
    MAPI_CONTACT_USER_FIELD4: "MAPI_CONTACT_USER_FIELD4",
    MAPI_DISTRIBUTION_LIST_NAME: "MAPI_DISTRIBUTION_LIST_NAME",
    MAPI_DISTRIBUTION_LIST_ONE_OFF_MEMBERS: "MAPI_DISTRIBUTION_LIST_ONE_OFF_MEMBERS",
    MAPI_DISTRIBUTION_LIST_MEMBERS: "MAPI_DISTRIBUTION_LIST_MEMBERS",
    MAPI_INSTANT_MESSAGING_ADDRESS: "MAPI_INSTANT_MESSAGING_ADDRESS",
    MAPI_DISTRIBUTION_LIST_STREAM: "MAPI_DISTRIBUTION_LIST_STREAM",
    MAPI_EMAIL_DISPLAY_NAME: "MAPI_EMAIL_DISPLAY_NAME",
    MAPI_EMAIL_ADDR_TYPE: "MAPI_EMAIL_ADDR_TYPE",
    MAPI_EMAIL_EMAIL_ADDRESS: "MAPI_EMAIL_EMAIL_ADDRESS",
    MAPI_EMAIL_ORIGINAL_DISPLAY_NAME: "MAPI_EMAIL_ORIGINAL_DISPLAY_NAME",
    MAPI_EMAIL1ORIGINAL_ENTRY_ID: "MAPI_EMAIL1ORIGINAL_ENTRY_ID",
    MAPI_EMAIL1RICH_TEXT_FORMAT: "MAPI_EMAIL1RICH_TEXT_FORMAT",
    MAPI_EMAIL1EMAIL_TYPE: "MAPI_EMAIL1EMAIL_TYPE",
    MAPI_EMAIL2DISPLAY_NAME: "MAPI_EMAIL2DISPLAY_NAME",
    MAPI_EMAIL2ENTRY_ID: "MAPI_EMAIL2ENTRY_ID",
    MAPI_EMAIL2ADDR_TYPE: "MAPI_EMAIL2ADDR_TYPE",
    MAPI_EMAIL2EMAIL_ADDRESS: "MAPI_EMAIL2EMAIL_ADDRESS",
    MAPI_EMAIL2ORIGINAL_DISPLAY_NAME: "MAPI_EMAIL2ORIGINAL_DISPLAY_NAME",
    MAPI_EMAIL2ORIGINAL_ENTRY_ID: "MAPI_EMAIL2ORIGINAL_ENTRY_ID",
    MAPI_EMAIL2RICH_TEXT_FORMAT: "MAPI_EMAIL2RICH_TEXT_FORMAT",
    MAPI_EMAIL3DISPLAY_NAME: "MAPI_EMAIL3DISPLAY_NAME",
    MAPI_EMAIL3ENTRY_ID: "MAPI_EMAIL3ENTRY_ID",
    MAPI_EMAIL3ADDR_TYPE: "MAPI_EMAIL3ADDR_TYPE",
    MAPI_EMAIL3EMAIL_ADDRESS: "MAPI_EMAIL3EMAIL_ADDRESS",
    MAPI_EMAIL3ORIGINAL_DISPLAY_NAME: "MAPI_EMAIL3ORIGINAL_DISPLAY_NAME",
    MAPI_EMAIL3ORIGINAL_ENTRY_ID: "MAPI_EMAIL3ORIGINAL_ENTRY_ID",
    MAPI_EMAIL3RICH_TEXT_FORMAT: "MAPI_EMAIL3RICH_TEXT_FORMAT",
    MAPI_FAX1ADDRESS_TYPE: "MAPI_FAX1ADDRESS_TYPE",
    MAPI_FAX1EMAIL_ADDRESS: "MAPI_FAX1EMAIL_ADDRESS",
    MAPI_FAX1ORIGINAL_DISPLAY_NAME: "MAPI_FAX1ORIGINAL_DISPLAY_NAME",
    MAPI_FAX1ORIGINAL_ENTRY_ID: "MAPI_FAX1ORIGINAL_ENTRY_ID",
    MAPI_FAX2ADDRESS_TYPE: "MAPI_FAX2ADDRESS_TYPE",
    MAPI_FAX2EMAIL_ADDRESS: "MAPI_FAX2EMAIL_ADDRESS",
    MAPI_FAX2ORIGINAL_DISPLAY_NAME: "MAPI_FAX2ORIGINAL_DISPLAY_NAME",
    MAPI_FAX2ORIGINAL_ENTRY_ID: "MAPI_FAX2ORIGINAL_ENTRY_ID",
    MAPI_FAX3ADDRESS_TYPE: "MAPI_FAX3ADDRESS_TYPE",
    MAPI_FAX3EMAIL_ADDRESS: "MAPI_FAX3EMAIL_ADDRESS",
    MAPI_FAX3ORIGINAL_DISPLAY_NAME: "MAPI_FAX3ORIGINAL_DISPLAY_NAME",
    MAPI_FAX3ORIGINAL_ENTRY_ID: "MAPI_FAX3ORIGINAL_ENTRY_ID",
    MAPI_FREE_BUSY_LOCATION: "MAPI_FREE_BUSY_LOCATION",
    MAPI_HOME_ADDRESS_COUNTRY_CODE: "MAPI_HOME_ADDRESS_COUNTRY_CODE",
    MAPI_WORK_ADDRESS_COUNTRY_CODE: "MAPI_WORK_ADDRESS_COUNTRY_CODE",
    MAPI_OTHER_ADDRESS_COUNTRY_CODE: "MAPI_OTHER_ADDRESS_COUNTRY_CODE",
    MAPI_ADDRESS_COUNTRY_CODE: "MAPI_ADDRESS_COUNTRY_CODE",
    MAPI_BIRTHDAY_LOCAL: "MAPI_BIRTHDAY_LOCAL",
    MAPI_WEDDING_ANNIVERSARY_LOCAL: "MAPI_WEDDING_ANNIVERSARY_LOCAL",
    MAPI_TASK_STATUS: "MAPI_TASK_STATUS",
    MAPI_TASK_START_DATE: "MAPI_TASK_START_DATE",
    MAPI_TASK_DUE_DATE: "MAPI_TASK_DUE_DATE",
    MAPI_TASK_ACTUAL_EFFORT: "MAPI_TASK_ACTUAL_EFFORT",
    MAPI_TASK_ESTIMATED_EFFORT: "MAPI_TASK_ESTIMATED_EFFORT",
    MAPI_TASK_FRECUR: "MAPI_TASK_FRECUR",
    MAPI_SEND_MEETING_AS_ICAL: "MAPI_SEND_MEETING_AS_ICAL",
    MAPI_APPOINTMENT_SEQUENCE: "MAPI_APPOINTMENT_SEQUENCE",
    MAPI_APPOINTMENT_SEQUENCE_TIME: "MAPI_APPOINTMENT_SEQUENCE_TIME",
    MAPI_APPOINTMENT_LAST_SEQUENCE: "MAPI_APPOINTMENT_LAST_SEQUENCE",
    MAPI_CHANGE_HIGHLIGHT: "MAPI_CHANGE_HIGHLIGHT",
    MAPI_BUSY_STATUS: "MAPI_BUSY_STATUS",
    MAPI_FEXCEPTIONAL_BODY: "MAPI_FEXCEPTIONAL_BODY",
    MAPI_APPOINTMENT_AUXILIARY_FLAGS: "MAPI_APPOINTMENT_AUXILIARY_FLAGS",
    MAPI_OUTLOOK_LOCATION: "MAPI_OUTLOOK_LOCATION",
    MAPI_MEETING_WORKSPACE_URL: "MAPI_MEETING_WORKSPACE_URL",
    MAPI_FORWARD_INSTANCE: "MAPI_FORWARD_INSTANCE",
    MAPI_LINKED_TASK_ITEMS: "MAPI_LINKED_TASK_ITEMS",
    MAPI_APPT_START_WHOLE: "MAPI_APPT_START_WHOLE",
    MAPI_APPT_END_WHOLE: "MAPI_APPT_END_WHOLE",
    MAPI_APPOINTMENT_START_TIME: "MAPI_APPOINTMENT_START_TIME",
    MAPI_APPOINTMENT_END_TIME: "MAPI_APPOINTMENT_END_TIME",
    MAPI_APPOINTMENT_END_DATE: "MAPI_APPOINTMENT_END_DATE",
    MAPI_APPOINTMENT_START_DATE: "MAPI_APPOINTMENT_START_DATE",
    MAPI_APPT_DURATION: "MAPI_APPT_DURATION",
    MAPI_APPOINTMENT_COLOR: "MAPI_APPOINTMENT_COLOR",
    MAPI_APPOINTMENT_SUB_TYPE: "MAPI_APPOINTMENT_SUB_TYPE",
    MAPI_APPOINTMENT_RECUR: "MAPI_APPOINTMENT_RECUR",
    MAPI_APPOINTMENT_STATE_FLAGS: "MAPI_APPOINTMENT_STATE_FLAGS",
    MAPI_RESPONSE_STATUS: "MAPI_RESPONSE_STATUS",
    MAPI_APPOINTMENT_REPLY_TIME: "MAPI_APPOINTMENT_REPLY_TIME",
    MAPI_RECURRING: "MAPI_RECURRING",
    MAPI_INTENDED_BUSY_STATUS: "MAPI_INTENDED_BUSY_STATUS",
    MAPI_APPOINTMENT_UPDATE_TIME: "MAPI_APPOINTMENT_UPDATE_TIME",
    MAPI_EXCEPTION_REPLACE_TIME: "MAPI_EXCEPTION_REPLACE_TIME",
    MAPI_OWNER_NAME: "MAPI_OWNER_NAME",
    MAPI_APPOINTMENT_REPLY_NAME: "MAPI_APPOINTMENT_REPLY_NAME",
    MAPI_RECURRENCE_TYPE: "MAPI_RECURRENCE_TYPE",
    MAPI_RECURRENCE_PATTERN: "MAPI_RECURRENCE_PATTERN",
    MAPI_TIME_ZONE_STRUCT: "MAPI_TIME_ZONE_STRUCT",
    MAPI_TIME_ZONE_DESCRIPTION: "MAPI_TIME_ZONE_DESCRIPTION",
    MAPI_CLIP_START: "MAPI_CLIP_START",
    MAPI_CLIP_END: "MAPI_CLIP_END",
    MAPI_ORIGINAL_STORE_ENTRY_ID: "MAPI_ORIGINAL_STORE_ENTRY_ID",
    MAPI_ALL_ATTENDEES_STRING: "MAPI_ALL_ATTENDEES_STRING",
    MAPI_AUTO_FILL_LOCATION: "MAPI_AUTO_FILL_LOCATION",
    MAPI_TO_ATTENDEES_STRING: "MAPI_TO_ATTENDEES_STRING",
    MAPI_CCATTENDEES_STRING: "MAPI_CCATTENDEES_STRING",
    MAPI_CONF_CHECK: "MAPI_CONF_CHECK",
    MAPI_CONFERENCING_TYPE: "MAPI_CONFERENCING_TYPE",
    MAPI_DIRECTORY: "MAPI_DIRECTORY",
    MAPI_ORGANIZER_ALIAS: "MAPI_ORGANIZER_ALIAS",
    MAPI_AUTO_START_CHECK: "MAPI_AUTO_START_CHECK",
    MAPI_AUTO_START_WHEN: "MAPI_AUTO_START_WHEN",
    MAPI_ALLOW_EXTERNAL_CHECK: "MAPI_ALLOW_EXTERNAL_CHECK",
    MAPI_COLLABORATE_DOC: "MAPI_COLLABORATE_DOC",
    MAPI_NET_SHOW_URL: "MAPI_NET_SHOW_URL",
    MAPI_ONLINE_PASSWORD: "MAPI_ONLINE_PASSWORD",
    MAPI_APPOINTMENT_PROPOSED_DURATION: "MAPI_APPOINTMENT_PROPOSED_DURATION",
    MAPI_APPT_COUNTER_PROPOSAL: "MAPI_APPT_COUNTER_PROPOSAL",
    MAPI_APPOINTMENT_PROPOSAL_NUMBER: "MAPI_APPOINTMENT_PROPOSAL_NUMBER",
    MAPI_APPOINTMENT_NOT_ALLOW_PROPOSE: "MAPI_APPOINTMENT_NOT_ALLOW_PROPOSE",
    MAPI_APPT_TZDEF_START_DISPLAY: "MAPI_APPT_TZDEF_START_DISPLAY",
    MAPI_APPT_TZDEF_END_DISPLAY: "MAPI_APPT_TZDEF_END_DISPLAY",
    MAPI_APPT_TZDEF_RECUR: "MAPI_APPT_TZDEF_RECUR",
    MAPI_REMINDER_MINUTES_BEFORE_START: "MAPI_REMINDER_MINUTES_BEFORE_START",
    MAPI_REMINDER_TIME: "MAPI_REMINDER_TIME",
    MAPI_REMINDER_SET: "MAPI_REMINDER_SET",
    MAPI_PRIVATE: "MAPI_PRIVATE",
    MAPI_AGING_DONT_AGE_ME: "MAPI_AGING_DONT_AGE_ME",
    MAPI_FORM_STORAGE: "MAPI_FORM_STORAGE",
    MAPI_SIDE_EFFECTS: "MAPI_SIDE_EFFECTS",
    MAPI_REMOTE_STATUS: "MAPI_REMOTE_STATUS",
    MAPI_PAGE_DIR_STREAM: "MAPI_PAGE_DIR_STREAM",
    MAPI_SMART_NO_ATTACH: "MAPI_SMART_NO_ATTACH",
    MAPI_COMMON_START: "MAPI_COMMON_START",
    MAPI_COMMON_END: "MAPI_COMMON_END",
    MAPI_TASK_MODE: "MAPI_TASK_MODE",
    MAPI_FORM_PROP_STREAM: "MAPI_FORM_PROP_STREAM",
    MAPI_REQUEST: "MAPI_REQUEST",
    MAPI_NON_SENDABLE_TO: "MAPI_NON_SENDABLE_TO",
    MAPI_NON_SENDABLE_CC: "MAPI_NON_SENDABLE_CC",
    MAPI_NON_SENDABLE_BCC: "MAPI_NON_SENDABLE_BCC",
    MAPI_COMPANIES: "MAPI_COMPANIES",
    MAPI_CONTACTS: "MAPI_CONTACTS",
    MAPI_PROP_DEF_STREAM: "MAPI_PROP_DEF_STREAM",
    MAPI_SCRIPT_STREAM: "MAPI_SCRIPT_STREAM",
    MAPI_CUSTOM_FLAG: "MAPI_CUSTOM_FLAG",
    MAPI_OUTLOOK_CURRENT_VERSION: "MAPI_OUTLOOK_CURRENT_VERSION",
    MAPI_CURRENT_VERSION_NAME: "MAPI_CURRENT_VERSION_NAME",
    MAPI_REMINDER_NEXT_TIME: "MAPI_REMINDER_NEXT_TIME",
    MAPI_HEADER_ITEM: "MAPI_HEADER_ITEM",
    MAPI_USE_TNEF: "MAPI_USE_TNEF",
    MAPI_TO_DO_TITLE: "MAPI_TO_DO_TITLE",
    MAPI_VALID_FLAG_STRING_PROOF: "MAPI_VALID_FLAG_STRING_PROOF",
    MAPI_LOG_TYPE: "MAPI_LOG_TYPE",
    MAPI_LOG_START: "MAPI_LOG_START",
    MAPI_LOG_DURATION: "MAPI_LOG_DURATION",
    MAPI_LOG_END: "MAPI_LOG_END",
}
