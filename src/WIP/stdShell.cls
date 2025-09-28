VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Issues:
' WIN32 MIGRATION
' - [ ] On windows, migrate to Win32 API rather than native calls.
' - [ ] File hashing should use Win32 API rather than .NET COM classes. Performance much better with Win32 API.
' - [ ] ReadText uses Input$ (locale/ANSI), no UTF-8/UTF-16 handling.
'FILE ENUMERATION
' - [ ] No binary read/write methods; no large-file chunking helpers.
'GENERAL
' - [ ] Support alternate data streams.`EShellFileTypeAlternateDataStream`
' - [ ] Path normalisation `\` and `/` should be interchangeable.
' - [ ] Better error handling generally, with descriptive errors.
' - [ ] stdShell#SecureDelete
' - [ ] stdsShell#Watch - Watch for changes to a file / folder


Public Enum stdShellKind
    EShellFileTypeUnknown
    EShellFileTypeRoot
    EShellFileTypeFile
    EShellFileTypeFolder
End Enum

Public Enum stdShellHashType
    EShellHashTypeMD5
    EShellHashTypeSHA1
    EShellHashTypeSHA256
End Enum

Public Enum stdShellExtendedProps
    [_stdShellExtendedPropsStart] = 0
    SystemAcquisitionID = [_stdShellExtendedPropsStart]
    SystemApplicationDefinedProperties
    SystemApplicationName
    SystemAppZoneIdentifier
    SystemAuthor
    SystemCachedFileUpdaterContentIdForConflictResolution
    SystemCachedFileUpdaterContentIdForStream
    SystemCapacity
    SystemCategory
    SystemComment
    SystemCompany
    SystemComputerName
    SystemContainedItems
    SystemContentStatus
    SystemContentType
    SystemCopyright
    SystemCreatorAppId
    SystemCreatorOpenWithUIOptions
    SystemDataObjectFormat
    SystemDateAccessed
    SystemDateAcquired
    SystemDateArchived
    SystemDateCompleted
    SystemDateCreated
    SystemDateImported
    SystemDateModified
    SystemDefaultSaveLocationDisplay
    SystemDueDate
    SystemEndDate
    SystemExpandoProperties
    SystemFileAllocationSize
    SystemFileAttributes
    SystemFileCount
    SystemFileDescription
    SystemFileExtension
    SystemFileFRN
    SystemFilename
    SystemFileOfflineAvailabilityStatus
    SystemFileOwner
    SystemFilePlaceholderStatus
    SystemFileVersion
    SystemFindData
    SystemFlagColor
    SystemFlagColorText
    SystemFlagStatus
    SystemFlagStatusText
    SystemFolderKind
    SystemFolderNameDisplay
    SystemFreeSpace
    SystemFullText
    SystemHighKeywords
    SystemImageParsingName
    SystemImportance
    SystemImportanceText
    SystemIsAttachment
    SystemIsDefaultNonOwnerSaveLocation
    SystemIsDefaultSaveLocation
    SystemIsDeleted
    SystemIsEncrypted
    SystemIsFlagged
    SystemIsFlaggedComplete
    SystemIsIncomplete
    SystemIsLocationSupported
    SystemIsPinnedToNameSpaceTree
    SystemIsRead
    SystemIsSearchOnlyItem
    SystemIsSendToTarget
    SystemIsShared
    SystemItemAuthors
    SystemItemClassType
    SystemItemDate
    SystemItemFolderNameDisplay
    SystemItemFolderPathDisplay
    SystemItemFolderPathDisplayNarrow
    SystemItemName
    SystemItemNameDisplay
    SystemItemNameDisplayWithoutExtension
    SystemItemNamePrefix
    SystemItemNameSortOverride
    SystemItemParticipants
    SystemItemPathDisplay
    SystemItemPathDisplayNarrow
    SystemItemSubType
    SystemItemType
    SystemItemTypeText
    SystemItemUrl
    SystemKEYWORDS
    SystemKind
    SystemKindText
    Systemlanguage
    SystemLastSyncError
    SystemLastWriterPackageFamilyName
    SystemLowKeywords
    SystemMediumKeywords
    SystemMileageInformation
    SystemMIMEType
    [_Null]
    SystemOfflineAvailability
    SystemOfflineStatus
    SystemOriginalFileName
    SystemOwnerSID
    SystemParentalRating
    SystemParentalRatingReason
    SystemParentalRatingsOrganization
    SystemParsingBindContext
    SystemParsingName
    SystemParsingPath
    SystemPerceivedType
    SystemPercentFull
    SystemPriority
    SystemPriorityText
    SystemProject
    SystemProviderItemID
    SystemRating
    SystemRatingText
    SystemRemoteConflictingFile
    SystemSensitivity
    SystemSensitivityText
    SystemSFGAOFlags
    SystemSharedWith
    SystemShareUserRating
    SystemSharingStatus
    Systemshell
    SystemSimpleRating
    SystemSize
    SystemSoftwareUsed
    SystemSourceItem
    SystemSourcePackageFamilyName
    SystemStartDate
    SystemStatus
    SystemStorageProviderCallerVersionInformation
    SystemStorageProviderError
    SystemStorageProviderFileChecksum
    SystemStorageProviderFileIdentifier
    SystemStorageProviderFileRemoteUri
    SystemStorageProviderFileVersion
    SystemStorageProviderFileVersionWaterline
    SystemStorageProviderId
    SystemStorageProviderShareStatuses
    SystemStorageProviderSharingStatus
    SystemStorageProviderStatus
    SystemSubject
    SystemSyncTransferStatus
    SystemThumbnail
    SystemThumbnailCacheId
    SystemThumbnailStream
    SystemTitle
    SystemTitleSortOverride
    SystemTotalFileSize
    SystemTrademarks
    SystemTransferOrder
    SystemTransferPosition
    SystemTransferSize
    SystemVolumeId
    SystemZoneIdentifier
    DocumentByteCount
    DocumentCharacterCount
    DocumentClientID
    DocumentContributor
    DocumentDateCreated
    DocumentDatePrinted
    DocumentDateSaved
    DocumentDivision
    DocumentDocumentID
    DocumentHiddenSlideCount
    DocumentLastAuthor
    DocumentLineCount
    DocumentManager
    DocumentMultimediaClipCount
    DocumentNoteCount
    DocumentPageCount
    DocumentParagraphCount
    DocumentPresentationFormat
    DocumentRevisionNumber
    DocumentSecurity
    DocumentSlideCount
    DocumentTEMPLATE
    DocumentTotalEditingTime
    DocumentVersion
    DocumentWordCount
    VideoCompression
    VideoDirector
    VideoEncodingBitrate
    VideoFourCC
    VideoFrameHeight
    VideoFrameRate
    VideoFrameWidth
    VideoHorizontalAspectRatio
    VideoIsSpherical
    VideoIsStereo
    VideoOrientation
    VideoSampleSize
    VideoStreamName
    VideoStreamNumber
    VideoTotalBitrate
    VideoTranscodedForSync
    VideoVerticalAspectRatio
    PhotoAperture
    PhotoApertureDenominator
    PhotoApertureNumerator
    PhotoBrightness
    PhotoBrightnessDenominator
    PhotoBrightnessNumerator
    PhotoCameraManufacturer
    PhotoCameraModel
    PhotoCameraSerialNumber
    PhotoContrast
    PhotoContrastText
    PhotoDateTaken
    PhotoDigitalZoom
    PhotoDigitalZoomDenominator
    PhotoDigitalZoomNumerator
    [_Event]
    PhotoEXIFVersion
    PhotoExposureBias
    PhotoExposureBiasDenominator
    PhotoExposureBiasNumerator
    PhotoExposureIndex
    PhotoExposureIndexDenominator
    PhotoExposureIndexNumerator
    PhotoExposureProgram
    PhotoExposureProgramText
    PhotoExposureTime
    PhotoExposureTimeDenominator
    PhotoExposureTimeNumerator
    PhotoFlash
    PhotoFlashEnergy
    PhotoFlashEnergyDenominator
    PhotoFlashEnergyNumerator
    PhotoFlashManufacturer
    PhotoFlashModel
    PhotoFlashText
    PhotoFNumber
    PhotoFNumberDenominator
    PhotoFNumberNumerator
    PhotoFocalLength
    PhotoFocalLengthDenominator
    PhotoFocalLengthInFilm
    PhotoFocalLengthNumerator
    PhotoFocalPlaneXResolution
    PhotoFocalPlaneXResolutionDenominator
    PhotoFocalPlaneXResolutionNumerator
    PhotoFocalPlaneYResolution
    PhotoFocalPlaneYResolutionDenominator
    PhotoFocalPlaneYResolutionNumerator
    PhotoGainControl
    PhotoGainControlDenominator
    PhotoGainControlNumerator
    PhotoGainControlText
    PhotoISOSpeed
    PhotoLensManufacturer
    PhotoLensModel
    PhotoLightSource
    PhotoMakerNote
    PhotoMakerNoteOffset
    PhotoMaxAperture
    PhotoMaxApertureDenominator
    PhotoMaxApertureNumerator
    PhotoMeteringMode
    PhotoMeteringModeText
    PhotoOrientation
    PhotoOrientationText
    PhotoPeopleNames
    PhotoPhotometricInterpretation
    PhotoPhotometricInterpretationText
    PhotoProgramMode
    PhotoProgramModeText
    PhotoRelatedSoundFile
    PhotoSaturation
    PhotoSaturationText
    PhotoSharpness
    PhotoSharpnessText
    PhotoShutterSpeed
    PhotoShutterSpeedDenominator
    PhotoShutterSpeedNumerator
    PhotoSubjectDistance
    PhotoSubjectDistanceDenominator
    PhotoSubjectDistanceNumerator
    PhotoTagViewAggregate
    PhotoTranscodedForSync
    PhotoWhiteBalance
    PhotoWhiteBalanceText
    MusicAlbumArtist
    MusicAlbumArtistSortOverride
    MusicAlbumID
    MusicAlbumTitle
    MusicAlbumTitleSortOverride
    MusicArtist
    MusicArtistSortOverride
    MusicBeatsPerMinute
    MusicComposer
    MusicComposerSortOverride
    MusicConductor
    MusicContentGroupDescription
    MusicDiscNumber
    MusicDisplayArtist
    MusicGenre
    MusicInitialKey
    MusicIsCompilation
    MusicLyrics
    MusicMood
    MusicPartOfSet
    Musicperiod
    MusicSynchronizedLyrics
    MusicTrackNumber
    MailAttachmentContents
    MailAttachmentNames
    MailBccAddress
    MailBccName
    MailCcAddress
    MailCcName
    MailConversationID
    MailConversationIndex
    MailDateReceived
    MailDateSent
    MailFlags
    MailFromAddress
    MailFromName
    MailHasAttachments
    MailIsFwdOrReply
    MailMessageClass
    MailParticipants
    MailProofInProgress
    MailSenderAddress
    MailSenderName
    MailStore
    MailToAddress
    MailToDoFlags
    MailToDoTitle
    MailToName
    MediaAuthorUrl
    MediaAverageLevel
    MediaClassPrimaryID
    MediaClassSecondaryID
    MediaCollectionGroupID
    MediaCollectionID
    MediaContentDistributor
    MediaContentID
    MediaCreatorApplication
    MediaCreatorApplicationVersion
    MediaDateEncoded
    MediaDateReleased
    MediaDlnaProfileID
    Mediaduration
    MediaDVDID
    MediaEncodedBy
    MediaEncodingSettings
    MediaEpisodeNumber
    MediaFrameCount
    MediaMCDI
    MediaMetadataContentProvider
    MediaProducer
    MediaPromotionUrl
    MediaProtectionType
    MediaProviderRating
    MediaProviderStyle
    MediaPublisher
    MediaSeasonNumber
    MediaSeriesName
    MediaSubscriptionContentId
    MediaSubTitle
    MediaThumbnailLargePath
    MediaThumbnailLargeUri
    MediaThumbnailSmallPath
    MediaThumbnailSmallUri
    MediaUniqueFileIdentifier
    MediaUserNoAutoInfo
    MediaUserWebUrl
    MediaWriter
    MediaYear
    ImageBitDepth
    ImageColorSpace
    ImageCompressedBitsPerPixel
    ImageCompressedBitsPerPixelDenominator
    ImageCompressedBitsPerPixelNumerator
    ImageCompression
    ImageCompressionText
    ImageDimensions
    ImageHorizontalResolution
    ImageHorizontalSize
    ImageImageID
    ImageResolutionUnit
    ImageVerticalResolution
    ImageVerticalSize

    [_stdShellExtendedPropsEnd]
End Enum

'source: https://github.com/KallunWillock/vbaWIA/blob/main/Properties/modWIA_EXIFMetadata.bas
Enum EShellExifProperties
    EXIFImageDateTimeOriginal = 36867
    EXIFImageTitle = 40091
    EXIFImageComments = 40092
    EXIFImageAuthor = 40093
    EXIFImageKeywords = 40094
    EXIFImageSubject = 40095
    EXIFGPSVer = 0                                        ' Version of the Global Positioning Systems (GPS) IFD, given as 2.0.0.0. This tag is mandatory when the GpsIFD tag is present. When the version is 2.0.0.0, the tag value is 0x02000000.
    EXIFGPSLatitudeRef = 1                                ' Null-terminated character string that specifies whether the latitude is north or south.: N: specifies north latitude, and: S: specifies south latitude.
    EXIFGPSLatitude = 2                                   ' Latitude. Latitude is expressed as three rational values giving the degrees, minutes, and seconds respectively. When degrees, minutes, and seconds are expressed, the format is dd/1, mm/1, ss/1. When degrees and minutes are used and, for example, fractions of minutes are given up to two decimal places, the format is dd/1, mmmm/100, 0/1.
    EXIFGPSLongitudeRef = 3                               ' Null-terminated character string that specifies whether the longitude is east or west longitude.: E: specifies east longitude, and: W: specifies west longitude.
    EXIFGPSLongitude = 4                                  ' Longitude. Longitude is expressed as three rational values giving the degrees, minutes, and seconds respectively. When degrees, minutes and seconds are expressed, the format is ddd/1, mm/1, ss/1. When degrees and minutes are used and, for example, fractions of minutes are given up to two decimal places, the format is ddd/1, mmmm/100, 0/1.
    EXIFGPSAltitudeRef = 5                                ' Reference altitude, in meters.
    EXIFGPSAltitude = 6                                   ' Altitude, in meters, based on the reference altitude specified by GpsAltitudeRef.
    EXIFGPSGPSTime = 7                                    ' Time as Coordinated Universal Time (UTC). The value is expressed as three rational numbers that give the hour, minute, and second.
    EXIFGPSGPSSatellites = 8                              ' Null-terminated character string that specifies the GPS satellites used for measurements. This tag can be used to specify the ID number, angle of elevation, azimuth, SNR, and other information about each satellite. The format is not specified. If the GPS receiver is incapable of taking measurements, the value of the tag must be set to: NULL.
    EXIFGPSGPSStatus = 9                                  ' Null-terminated character string that specifies the status of the GPS receiver when the image is recorded.: A: means measurement is in progress, and: V: means the measurement is Interoperability.
    EXIFGPSGPSMeasureMode = 10                            ' Null-terminated character string that specifies the GPS measurement mode.: 2: specifies 2-D measurement, and: 3: specifies 3-D measurement.
    EXIFGPSGPSDop = 11                                    ' GPS DOP (data degree of precision). An HDOP value is written during 2-D measurement, and a PDOP value is written during 3-D measurement.
    EXIFGPSSpeedRef = 12                                  ' Null-terminated character string that specifies the unit used to express the GPS receiver speed of movement.: K,: M, and: N: represent kilometers per hour, miles per hour, and knots respectively.
    EXIFGPSSpeed = 13                                     ' Speed of the GPS receiver movement.
    EXIFGPSTrackRef = 14                                  ' Null-terminated character string that specifies the reference for giving the direction of GPS receiver movement.: T: specifies true direction, and: M: specifies magnetic direction.
    EXIFGPSTrack = 15                                     ' Direction of GPS receiver movement. The range of values is from 0.00 to 359.99.
    EXIFGPSImgDirRef = 16                                 ' Null-terminated character string that specifies the reference for the direction of the image when it is captured.: T: specifies true direction, and: M: specifies magnetic direction.
    EXIFGPSImgDir = 17                                    ' Direction of the image when it was captured. The range of values is from 0.00 to 359.99.
    EXIFGPSMapDatum = 18                                  ' Null-terminated character string that specifies geodetic survey data used by the GPS receiver. If the survey data is restricted to Japan, the value of this tag is: TOKYO: or: WGS-84.
    EXIFGPSDestLatRef = 19                                ' Null-terminated character string that specifies whether the latitude of the destination point is north or south latitude.: N: specifies north latitude, and: S: specifies south latitude.
    EXIFGPSDestLat = 20                                   ' Latitude of the destination point. The latitude is expressed as three rational values giving the degrees, minutes, and seconds respectively. When degrees, minutes, and seconds are expressed, the format is dd/1, mm/1, ss/1. When degrees and minutes are used and, for example, fractions of minutes are given up to two decimal places, the format is dd/1, mmmm/100, 0/1.
    EXIFGPSDestLongRef = 21                               ' Null-terminated character string that specifies whether the longitude of the destination point is east or west longitude.: E: specifies east longitude, and: W: specifies west longitude.
    EXIFGPSDestLong = 22                                  ' Longitude of the destination point. The longitude is expressed as three rational values giving the degrees, minutes, and seconds respectively. When degrees, minutes, and seconds are expressed, the format is ddd/1, mm/1, ss/1. When degrees and minutes are used and, for example, fractions of minutes are given up to two decimal places, the format is ddd/1, mmmm/100, 0/1.
    EXIFGPSDestBearRef = 23                               ' Null-terminated character string that specifies the reference used for giving the bearing to the destination point.: T: specifies true direction, and: M: specifies magnetic direction.
    EXIFGPSDestBear = 24                                  ' Bearing to the destination point. The range of values is from 0.00 to 359.99.
    EXIFGPSDestDistRef = 25                               ' Null-terminated character string that specifies the unit used to express the distance to the destination point. K, M, and N represent kilometers, miles, and knots respectively.
    EXIFGPSDestDist = 26                                  ' Distance to the destination point.
    EXIFNewSubfileType = 254                              ' Type of data in a subfile.
    EXIFSubfileType = 255                                 ' Type of data in a subfile.
    EXIFImageWidth = 256                                  ' Number of pixels per row.
    EXIFImageHeight = 257                                 ' Number of pixel rows.
    EXIFBitsPerSample = 258                               ' Number of bits per color component. See also: SamplesPerPixel.
    EXIFCompression = 259                                 ' Compression scheme used for the image data.
    EXIFPhotometricInterp = 262                           ' How pixel data will be interpreted.
    EXIFThreshHolding = 263                               ' Technique used to convert from gray pixels to black and white pixels.
    EXIFCellWidth = 264                                   ' Width of the dithering or halftoning matrix.
    EXIFCellHeight = 265                                  ' Height of the dithering or halftoning matrix.
    EXIFFillOrder = 266                                   ' Logical order of bits in a byte.
    EXIFDocumentName = 269                                ' Null-terminated character string that specifies the name of the document from which the image was scanned.
    EXIFImageDescription = 270                            ' Null-terminated character string that specifies the title of the image.
    EXIFEquipMake = 271                                   ' Null-terminated character string that specifies the manufacturer of the equipment used to record the image.
    EXIFEquipModel = 272                                  ' Null-terminated character string that specifies the model name or model number of the equipment used to record the image.
    EXIFStripOffsets = 273                                ' For each strip, the byte offset of that strip. See also: RowsPerStrip: and: StripBytesCount.
    EXIFOrientation = 274                                 ' Image orientation viewed in terms of rows and columns.
    EXIFSamplesPerPixel = 277                             ' Number of color components per pixel.
    EXIFRowsPerStrip = 278                                ' Number of rows per strip. See also: StripBytesCount: and: StripOffsets.
    EXIFStripBytesCount = 279                             ' For each strip, the total number of bytes in that strip.
    EXIFMinSampleValue = 280                              ' For each color component, the minimum value assigned to that component. See also: SamplesPerPixel.
    EXIFMaxSampleValue = 281                              ' For each color component, the maximum value assigned to that component. See also: SamplesPerPixel.
    EXIFXResolution = 282                                 ' Number of pixels per unit in the image width (x) direction. The unit is specified by: ResolutionUnit.
    EXIFYResolution = 283                                 ' Number of pixels per unit in the image height (y) direction. The unit is specified by: ResolutionUnit.
    EXIFPlanarConfig = 284                                ' Whether pixel components are recorded in chunky or planar format.
    EXIFPageName = 285                                    ' Null-terminated character string that specifies the name of the page from which the image was scanned.
    EXIFXPosition = 286                                   ' Offset from the left side of the page to the left side of the image. The unit of measure is specified by: ResolutionUnit.
    EXIFYPosition = 287                                   ' Offset from the top of the page to the top of the image. The unit of measure is specified by: ResolutionUnit.
    EXIFFreeOffset = 288                                  ' For each string of contiguous unused bytes, the byte offset of that string.
    EXIFFreeByteCounts = 289                              ' For each string of contiguous unused bytes, the number of bytes in that string.
    EXIFGrayResponseUnit = 290                            ' Precision of the number specified by GrayResponseCurve. 1 specifies tenths, 2 specifies hundredths, 3 specifies thousandths, and so on.
    EXIFGrayResponseCurve = 291                           ' For each possible pixel value in a grayscale image, the optical density of that pixel value.
    EXIFT4Option = 292                                    ' Set of flags that relate to T4 encoding.
    EXIFT6Option = 293                                    ' Set of flags that relate to T6 encoding.
    EXIFResolutionUnit = 296                              ' Unit of measure for the horizontal resolution and the vertical resolution.
    EXIFPageNumber = 297                                  ' Page number of the page from which the image was scanned.
    EXIFTransferFunction = 301                            ' Tables that specify transfer functions for the image.
    EXIFSoftwareUsed = 305                                ' Null-terminated character string that specifies the name and version of the software or firmware of the device used to generate the image.
    EXIFdateTime = 306                                    ' Date and time the image was created.
    EXIFArtist = 315                                      ' Null-terminated character string that specifies the name of the person who created the image.
    EXIFHostComputer = 316                                ' Null-terminated character string that specifies the computer and/or operating system used to create the image.
    EXIFPredictor = 317                                   ' Type of prediction scheme that was applied to the image data before the encoding scheme was applied.
    EXIFWhitePoint = 318                                  ' Chromaticity of the white point of the image.
    EXIFPrimaryChromaticities = 319                       ' For each of the three primary colors in the image, the chromaticity of that color.
    EXIFColorMap = 320                                    ' Color palette (lookup table) for a palette-indexed image.
    EXIFHalftoneHints = 321                               ' Information used by the halftone function
    EXIFTileWidth = 322                                   ' Number of pixel columns in each tile.
    EXIFTileLength = 323                                  ' Number of pixel rows in each tile.
    EXIFTileOffset = 324                                  ' For each tile, the byte offset of that tile.
    EXIFTileByteCounts = 325                              ' For each tile, the number of bytes in that tile.
    EXIFInkSet = 332                                      ' Set of inks used in a separated image.
    EXIFInkNames = 333                                    ' Sequence of concatenated, null-terminated, character strings that specify the names of the inks used in a separated image.
    EXIFNumberOfInks = 334                                ' Number of inks.
    EXIFDotRange = 336                                    ' Color component values that correspond to a 0 percent dot and a 100 percent dot.
    EXIFTargetPrinter = 337                               ' Null-terminated character string that describes the intended printing environment.
    EXIFExtraSamples = 338                                ' Number of extra color components. For example, one extra component might hold an alpha value.
    EXIFSampleFormat = 339                                ' For each color component, the numerical format (unsigned, signed, floating point) of that component. See also: SamplesPerPixel.
    EXIFSMinSampleValue = 342                             ' For each color component, the minimum value of that component. See also: SamplesPerPixel.
    EXIFSMaxSampleValue = 512                             ' For each color component, the maximum value of that component. See also: SamplesPerPixel.
    EXIFTransferRange = 513                               ' Table of values that extends the range of the transfer function.
    EXIFJPEGProc = 514                                    ' JPEG compression process.
    EXIFJPEGInterFormat = 515                             ' Offset to the start of a JPEG bitstream.
    EXIFJPEGInterLength = 517                             ' Length, in bytes, of the JPEG bitstream.
    EXIFJPEGRestartInterval = 518                         ' Length of the restart interval.
    EXIFJPEGLosslessPredictors = 519                      ' For each color component, a lossless predictor-selection value for that component. See also: SamplesPerPixel.
    EXIFJPEGPointTransforms = 520                         ' For each color component, a point transformation value for that component. See also: SamplesPerPixel.
    EXIFJPEGQTables = 521                                 ' For each color component, the offset to the quantization table for that component. See also: SamplesPerPixel.
    EXIFJPEGDCTables = 529                                ' For each color component, the offset to the DC Huffman table (or lossless Huffman table) for that component. See also: SamplesPerPixel.
    EXIFJPEGACTables = 530                                ' For each color component, the offset to the AC Huffman table for that component. See also: SamplesPerPixel.
    EXIFYCbCrCoefficients = 531                           ' Coefficients for transformation from RGB to YCbCr image data.
    EXIFYCbCrSubsampling = 532                            ' Sampling ratio of chrominance components in relation to the luminance component.
    EXIFYCbCrPositioning = 769                            ' Position of chrominance components in relation to the luminance component.
    EXIFREFBlackWhite = 770                               ' Reference black point value and reference white point value.
    EXIFGamma = 771                                       ' Gamma value attached to the image. The gamma value is stored as a rational number (pair of: long) with a numerator of 100000. For example, a gamma value of 2.2 is stored as the pair (100000, 45455).
    EXIFICCProfileDescriptor = 800                        ' Null-terminated character string that identifies an ICC profile.
    EXIFSRGBRenderingIntent = 20481                       ' Saturation intent, which is suitable for charts and graphs, preserves saturation at the expense of hue and lightness.
    EXIFImageTitle = 20482                                ' Null-terminated character string that specifies the title of the image.
    EXIFResolutionXUnit = 20483                           ' Units in which to display horizontal resolution.
    EXIFResolutionYUnit = 20484                           ' Units in which to display vertical resolution.
    EXIFResolutionXLengthUnit = 20485                     ' Units in which to display the image width.
    EXIFResolutionYLengthUnit = 20486                     ' Units in which to display the image height.
    EXIFPrintFlags = 20487                                ' Sequence of one-byte Boolean values that specify printing options.
    EXIFPrintFlagsVersion = 20488                         ' Print flags version.
    EXIFPrintFlagsCrop = 20489                            ' Print flags center crop marks.
    EXIFPrintFlagsBleedWidth = 20490                      ' Print flags bleed width.
    EXIFPrintFlagsBleedWidthScale = 20491                 ' Print flags bleed width scale.
    EXIFHalftoneLPI = 20492                               ' Ink's screen frequency, in lines per inch.
    EXIFHalftoneLPIUnit = 20493                           ' Units for the screen frequency.
    EXIFHalftoneDegree = 20494                            ' Angle for screen.
    EXIFHalftoneShape = 20495                             ' Shape of the halftone dots.
    EXIFHalftoneMisc = 20496                              ' Miscellaneous halftone information.
    EXIFHalftoneScreen = 20497                            ' Boolean value that specifies whether to use the printer's default screens.
    EXIFJPEGQuality = 20498                               ' Private tag used by the Adobe Photoshop format. Not for public use.
    EXIFGridSize = 20499                                  ' Block of information about grids and guides.
    EXIFThumbnailFormat = 20500                           ' Format of the thumbnail image.
    EXIFThumbnailWidth = 20501                            ' Width, in pixels, of the thumbnail image.
    EXIFThumbnailHeight = 20502                           ' Height, in pixels, of the thumbnail image.
    EXIFThumbnailColorDepth = 20503                       ' bits per pixel (BPP) for the thumbnail image.
    EXIFThumbnailPlanes = 20504                           ' Number of color planes for the thumbnail image.
    EXIFThumbnailRawBytes = 20505                         ' Byte offset between rows of pixel data.
    EXIFThumbnailSize = 20506                             ' Total size, in bytes, of the thumbnail image.
    EXIFThumbnailCompressedSize = 20507                   ' Compressed size, in bytes, of the thumbnail image.
    EXIFColorTransferFunction = 20512                     ' Table of values that specify color transfer functions.
    EXIFThumbnailData = 20513                             ' Raw thumbnail bits in JPEG or RGB format. Depends on ThumbnailFormat.
    EXIFThumbnailImageWidth = 20514                       ' Number of pixels per row in the thumbnail image.
    EXIFThumbnailImageHeight = 20515                      ' Number of pixel rows in the thumbnail image.
    EXIFThumbnailBitsPerSample = 20516                    ' Number of bits per color component in the thumbnail image. See also: ThumbnailSamplesPerPixel.
    EXIFThumbnailCompression = 20517                      ' Compression scheme used for thumbnail image data.
    EXIFThumbnailPhotometricInterp = 20518                ' How thumbnail pixel data will be interpreted.
    EXIFThumbnailImageDescription = 20519                 ' Null-terminated character string that specifies the title of the image.
    EXIFThumbnailEquipMake = 20520                        ' Null-terminated character string that specifies the manufacturer of the equipment used to record the thumbnail image.
    EXIFThumbnailEquipModel = 20521                       ' Null-terminated character string that specifies the model name or model number of the equipment used to record the thumbnail image.
    EXIFThumbnailStripOffsets = 20522                     ' For each strip in the thumbnail image, the byte offset of that strip. See also: ThumbnailRowsPerStrip: and: ThumbnailStripBytesCount.
    EXIFThumbnailOrientation = 20523                      ' Thumbnail image orientation in terms of rows and columns. See also: Orientation.
    EXIFThumbnailSamplesPerPixel = 20524                  ' Number of color components per pixel in the thumbnail image.
    EXIFThumbnailRowsPerStrip = 20525                     ' Number of rows per strip in the thumbnail image. See also: ThumbnailStripBytesCount: and: ThumbnailStripOffsets.
    EXIFThumbnailStripBytesCount = 20526                  ' For each thumbnail image strip, the total number of bytes in that strip.
    EXIFThumbnailResolutionX = 20527                      ' Thumbnail resolution in the width direction. The resolution unit is given in: ThumbnailResolutionUnit.
    EXIFThumbnailResolutionY = 20528                      ' Thumbnail resolution in the height direction. The resolution unit is given in: ThumbnailResolutionUnit.
    EXIFThumbnailPlanarConfig = 20529                     ' Whether pixel components in the thumbnail image are recorded in chunky or planar format. See also: PlanarConfig.
    EXIFThumbnailResolutionUnit = 20530                   ' Unit of measure for the horizontal resolution and the vertical resolution of the thumbnail image. See also: ResolutionUnit.
    EXIFThumbnailTransferFunction = 20531                 ' Tables that specify transfer functions for the thumbnail image. See also: TransferFunction.
    EXIFThumbnailSoftwareUsed = 20532                     ' Null-terminated character string that specifies the name and version of the software or firmware of the device used to generate the thumbnail image.
    EXIFThumbnailDateTime = 20533                         ' Date and time the thumbnail image was created. See also: DateTime.
    EXIFThumbnailArtist = 20534                           ' Null-terminated character string that specifies the name of the person who created the thumbnail image.
    EXIFThumbnailWhitePoint = 20535                       ' Chromaticity of the white point of the thumbnail image. See also: WhitePoint.
    EXIFThumbnailPrimaryChromaticities = 20536            ' For each of the three primary colors in the thumbnail image, the chromaticity of that color. See also: PrimaryChromaticities.
    EXIFThumbnailYCbCrCoefficients = 20537                ' Coefficients for transformation from RGB to YCbCr data for the thumbnail image. See also: YCbCrCoefficients.
    EXIFThumbnailYCbCrSubsampling = 20538                 ' Sampling ratio of chrominance components in relation to the luminance component for the thumbnail image. See also: YCbCrSubsampling.
    EXIFThumbnailYCbCrPositioning = 20539                 ' Position of chrominance components in relation to the luminance component for the thumbnail image. See also: YCbCrPositioning.
    EXIFThumbnailRefBlackWhite = 20624                    ' Reference black point value and reference white point value for the thumbnail image. See also: REFBlackWhite.
    EXIFThumbnailCopyRight = 20625                        ' Null-terminated character string that contains copyright information for the thumbnail image.
    EXIFLuminanceTable = 20736                            ' Luminance table. The luminance table and the chrominance table are used to control JPEG quality. A valid luminance or chrominance table has 64 entries of type TypeShort. If an image has either a luminance table or a chrominance table, then it must have both tables.
    EXIFChrominanceTable = 20737                          ' Chrominance table. The luminance table and the chrominance table are used to control JPEG quality. A valid luminance or chrominance table has 64 entries of type TypeShort. If an image has either a luminance table or a chrominance table, then it must have both tables.
    EXIFFrameDelay = 20738                                ' Time delay, in hundredths of a second, between two frames in an animated GIF image.
    EXIFLoopCount = 20739                                 ' For an animated GIF image, the number of times to display the animation. A value of 0 specifies that the animation should be displayed infinitely.
    EXIFGlobalPalette = 20740                             ' Color palette for an indexed bitmap in a GIF image.
    EXIFIndexBackground = 20752                           ' Index of the background color in the palette of a GIF image.
    EXIFIndexTransparent = 20753                          ' Index of the transparent color in the palette of a GIF image.
    EXIFPixelUnit = 20754                                 ' Unit for PixelPerUnitX and PixelPerUnitY.
    EXIFPixelPerUnitX = 20755                             ' Pixels per unit in the x direction.
    EXIFPixelPerUnitY = 33432                             ' Pixels per unit in the y direction.
    EXIFPaletteHistogram = 33434                          ' Palette histogram.
    EXIFCopyright = 33437                                 ' Null-terminated character string that contains copyright information.
    EXIFExposureTime = 34665                          ' Exposure time, measured in seconds.
    EXIFFNumber = 34675                               ' F number.
    EXIFIFD = 34850                                   ' Private tag used by GDI+. Not for public use. GDI+ uses this tag to locate Exif-specific information.
    EXIFICCProfile = 34852                                ' ICC profile embedded in the image.
    EXIFExposureProg = 34853                          ' Class of the program used by the camera to set exposure when the picture is taken.
    EXIFSpectralSense = 34855                         ' Null-terminated character string that specifies the spectral sensitivity of each channel of the camera used. The string is compatible with the standard developed by the ASTM Technical Committee.
    EXIFGPSIFD = 34856                                    ' Offset to a block of GPS property items. Property items whose tags have the prefix Gps are stored in the GPS block. The GPS property items are defined in the EXIF specification. GDI+ uses this tag to locate GPS information, but GDI+ does not expose this tag for public use.
    EXIFISOSpeed = 36864                              ' ISO speed and ISO latitude of the camera or input device as specified in ISO 12232.
    EXIFOECF = 36867                                  ' Optoelectronic conversion function (OECF) specified in ISO 14524. The OECF is the relationship between the camera optical input and the image values.
    EXIFVer = 36868                                   ' Version of the EXIF standard supported. Nonexistence of this field is taken to mean nonconformance to the standard. Conformance to the standard is indicated by recording 0210 as a 4-byte ASCII string. Because the type is TypeUndefined, there is no NULL terminator.
    EXIFDTOrig = 37121                                ' Date and time when the original image data was generated. For a DSC, the date and time when the picture was taken. The format is YYYY:MM:DD HH:MM:SS with time shown in 24-hour format and the date and time separated by one blank character (0x2000). The character string length is 20 bytes including the NULL terminator. When the field is empty, it is treated as unknown.
    EXIFDTDigitized = 37122                           ' The format is YYYY:MM:DD HH:MM:SS with time shown in 24-hour format and the date and time separated by one blank character (0x2000). The character string length is 20 bytes including the NULL terminator. When the field is empty, it is treated as unknown.
    EXIFCompConfig = 37377                            ' However, because PhotometricInterp can only express the order of Y, Cb, and Cr, this tag is provided for cases when compressed data uses components other than Y, Cb, and Cr and to support other sequences.
    EXIFCompBPP = 37378                               ' Information specific to compressed data. The compression mode used for a compressed image is indicated in unit BPP.
    EXIFShutterSpeed = 37379                          ' Shutter speed. The unit is the Additive System of Photographic Exposure (APEX) value.
    EXIFAperture = 37380                              ' Lens aperture. The unit is the APEX value.
    EXIFBrightness = 37381                            ' Brightness value. The unit is the APEX value. Ordinarily it is given in the range of -99.99 to 99.99.
    EXIFExposureBias = 37382                          ' Exposure bias. The unit is the APEX value. Ordinarily it is given in the range -99.99 to 99.99.
    EXIFMaxAperture = 37383                           ' Smallest F number of the lens. The unit is the APEX value. Ordinarily it is given in the range of 00.00 to 99.99, but it is not limited to this range.
    EXIFSubjectDist = 37384                           ' Distance to the subject, measured in meters.
    EXIFMeteringMode = 37385                          ' Metering mode.
    EXIFLightSource = 37386                           ' Type of light source.
    EXIFFlash = 37500                                 ' Flash status. This tag is recorded when an image is taken using a strobe light (flash). Bit 0 indicates the flash firing status, and bits 1 and 2 indicate the flash return status.
    EXIFFocalLength = 37510                           ' Actual focal length, in millimeters, of the lens. Conversion is not made to the focal length of a 35 millimeter film camera.
    EXIFMakerNote = 37520                             ' Note tag. A tag used by manufacturers of EXIF writers to record information. The contents are up to the manufacturer.
    EXIFUserComment = 37521                           ' Comment tag. A tag used by EXIF users to write keywords or comments about the image besides those in ImageDescription and without the character-code limitations of the ImageDescription tag.
    EXIFDTSubsec = 37522                              ' Null-terminated character string that specifies a fraction of a second for the DateTime tag.
    EXIFDTOrigSS = 40960                              ' Null-terminated character string that specifies a fraction of a second for the ExifDTOrig tag.
    EXIFDTDigSS = 40961                               ' Null-terminated character string that specifies a fraction of a second for the ExifDTDigitized tag.
    EXIFFPXVer = 40962                                ' FlashPix format version supported by an FPXR file. If the FPXR function supports FlashPix format version 1.0, this is indicated similarly to ExifVer by recording 0100 as a 4-byte ASCII string. Because the type is TypeUndefined, there is no NULL terminator.
    EXIFColorSpace = 40963                            ' Color space specifier. Normally sRGB (=1) is used to define the color space based on the PC monitor conditions and environment. If a color space other than sRGB is used, Uncalibrated (=0xFFFF) is set. Image data recorded as Uncalibrated can be treated as sRGB when it is converted to FlashPix.
    EXIFPixXDim = 40964                               ' Information specific to compressed data. When a compressed file is recorded, the valid width of the meaningful image must be recorded in this tag, whether or not there is padding data or a restart marker. This tag should not exist in an uncompressed file.
    EXIFPixYDim = 40965                               ' Information specific to compressed data. When a compressed file is recorded, the valid height of the meaningful image must be recorded in this tag whether or not there is padding data or a restart marker. This tag should not exist in an uncompressed file. Because data padding is unnecessary in the vertical direction, the number of lines recorded in this valid image height tag will be the same as that recorded in the SOF.
    EXIFRelatedWav = 41483                            ' The name of an audio file related to the image data. The only relational information recorded is the EXIF audio file name and extension (an ASCII string that consists of 8 characters plus a period (.), plus 3 characters). The path is not recorded. When you use this tag, audio files must be recorded in conformance with the EXIF audio format. Writers can also store audio data within APP2 as FlashPix extension stream data.
    EXIFInterop = 41484                               ' Offset to a block of property items that contain interoperability information.
    EXIFFlashEnergy = 41486                           ' Strobe energy, in Beam Candle Power Seconds (BCPS), at the time the image was captured.
    EXIFSpatialFR = 41487                             ' Camera or input device spatial frequency table and SFR values in the image width, image height, and diagonal direction, as specified in ISO 12233.
    EXIFFocalXRes = 41488                             ' Number of pixels in the image width (x) direction per unit on the camera focal plane. The unit is specified in ExifFocalResUnit.
    EXIFFocalYRes = 41492                             ' Number of pixels in the image height (y) direction per unit on the camera focal plane. The unit is specified in ExifFocalResUnit.
    EXIFFocalResUnit = 41493                          ' Unit of measure for ExifFocalXRes and ExifFocalYRes.
    EXIFSubjectLoc = 41495                            ' Location of the main subject in the scene. The value of this tag represents the pixel at the center of the main subject relative to the left edge. The first value indicates the column number, and the second value indicates the row number.
    EXIFExposureIndex = 41728                         ' Exposure index selected on the camera or input device at the time the image was captured.
    EXIFSensingMethod = 41729                         ' Image sensor type on the camera or input device.
    EXIFFileSource = 41730                            ' The image source. If a DSC recorded the image, the value of this tag is 3.
    EXIFIFDCustomRendered = 41985
    EXIFIFDExposureMode = 41986
    EXIFIFDWhiteBalance = 41987
    EXIFIFDDigitalZoomRatio = 41988
    EXIFIFDFocalLengthIn35mmFormat = 41989
    EXIFIFDSceneCaptureType = 41990
    EXIFIFDGainControl = 41991
    EXIFIFDContrast = 41992
    EXIFIFDSaturation = 41993
    EXIFIFDSharpness = 41994
    EXIFIFDDeviceSettingDescription = 41995
    EXIFIFDSubjectDistanceRange = 41996
    EXIFIFDImageUniqueID = 42016
    EXIFIFDOwnerName = 42032
    EXIFIFDSerialNumber = 42033
    EXIFIFDLensInfo = 42034
    EXIFIFDLensMake = 42035
    EXIFIFDLensModel = 42036
    EXIFIFDLensSerialNumber = 42037
    EXIFIFDCompositeImage = 42080
    EXIFIFDCompositeImageCount = 42081
    EXIFIFDCompositeImageExposureTimes = 42082
    EXIFIFDGDALMetadata = 42112
    EXIFIFDGDALNoData = 42113
    EXIFIFDGamma = 42240
    EXIFIFDExpandSoftware = 44992
    EXIFIFDExpandLens = 44993
    EXIFIFDExpandFilm = 44994
    EXIFIFDExpandFilterLens = 44995
    EXIFIFDExpandScanner = 44996
    EXIFIFDExpandFlashLamp = 44997
    EXIFIFDHasselbladRawImage = 46275
    EXIFIFDPixelFormat = 48129
    EXIFIFDTransformation = 48130
    EXIFIFDUncompressed = 48131
    EXIFIFDImageType = 48132
    EXIFIFDImageWidth = 48256
    EXIFIFDImageHeight = 48257
    EXIFIFDWidthResolution = 48258
    EXIFIFDHeightResolution = 48259
    EXIFIFDImageOffset = 48320
    EXIFIFDImageByteCount = 48321
    EXIFIFDAlphaOffset = 48322
    EXIFIFDAlphaByteCount = 48323
    EXIFIFDImageDataDiscard = 48324
    EXIFIFDAlphaDataDiscard = 48325
    EXIFIFDOceScanjobDesc = 50215
    EXIFIFDOceApplicationSelector = 50216
    EXIFIFDOceIDNumber = 50217
    EXIFIFDOceImageLogic = 50218
    EXIFIFDAnnotations = 50255
    EXIFIFDPrintIM = 50341
    EXIFIFDHasselbladExif = 50459
    EXIFIFDOriginalFileName = 50547
    EXIFIFDUSPTOOriginalContentType = 50560
    EXIFIFDCR2CFAPattern = 50656
    EXIFIFDDNGVersion = 50706
    EXIFIFDDNGBackwardVersion = 50707
    EXIFIFDUniqueCameraModel = 50708
    EXIFIFDLocalizedCameraModel = 50709
    EXIFIFDCFAPlaneColor = 50710
    EXIFIFDCFALayout = 50711
    EXIFIFDLinearizationTable = 50712
    EXIFIFDBlackLevelRepeatDim = 50713
    EXIFIFDBlackLevel = 50714
    EXIFIFDBlackLevelDeltaH = 50715
    EXIFIFDBlackLevelDeltaV = 50716
    EXIFIFDWhiteLevel = 50717
    EXIFIFDDefaultScale = 50718
    EXIFIFDDefaultCropOrigin = 50719
    EXIFIFDDefaultCropSize = 50720
    EXIFIFDColorMatrix1 = 50721
    EXIFIFDColorMatrix2 = 50722
    EXIFIFDCameraCalibration1 = 50723
    EXIFIFDCameraCalibration2 = 50724
    EXIFIFDReductionMatrix1 = 50725
    EXIFIFDReductionMatrix2 = 50726
    EXIFIFDAnalogBalance = 50727
    EXIFIFDAsShotNeutral = 50728
    EXIFIFDAsShotWhiteXY = 50729
    EXIFIFDBaselineExposure = 50730
    EXIFIFDBaselineNoise = 50731
    EXIFIFDBaselineSharpness = 50732
    EXIFIFDBayerGreenSplit = 50733
    EXIFIFDLinearResponseLimit = 50734
    EXIFIFDCameraSerialNumber = 50735
    EXIFIFDDNGLensInfo = 50736
    EXIFIFDChromaBlurRadius = 50737
    EXIFIFDAntiAliasStrength = 50738
    EXIFIFDShadowScale = 50739
    EXIFIFDSR2Private = 50740
    EXIFIFDMakerNoteSafety = 50741
    EXIFIFDRawImageSegmentation = 50752
    EXIFIFDCalibrationIlluminant1 = 50778
    EXIFIFDCalibrationIlluminant2 = 50779
    EXIFIFDBestQualityScale = 50780
    EXIFIFDRawDataUniqueID = 50781
    EXIFIFDAliasLayerMetadata = 50784
    EXIFIFDOriginalRawFileName = 50827
    EXIFIFDOriginalRawFileData = 50828
    EXIFIFDActiveArea = 50829
    EXIFIFDMaskedAreas = 50830
    EXIFIFDAsShotICCProfile = 50831
    EXIFIFDAsShotPreProfileMatrix = 50832
    EXIFIFDCurrentICCProfile = 50833
    EXIFIFDCurrentPreProfileMatrix = 50834
    EXIFIFDColorimetricReference = 50879
    EXIFIFDSRawType = 50885
    EXIFIFDPanasonicTitle = 50898
    EXIFIFDPanasonicTitle2 = 50899
    EXIFIFDCameraCalibrationSig = 50931
    EXIFIFDProfileCalibrationSig = 50932
    EXIFIFDProfileIFD = 50933
    EXIFIFDAsShotProfileName = 50934
    EXIFIFDNoiseReductionApplied = 50935
    EXIFIFDProfileName = 50936
    EXIFIFDProfileHueSatMapDims = 50937
    EXIFIFDProfileHueSatMapData1 = 50938
    EXIFIFDProfileHueSatMapData2 = 50939
    EXIFIFDProfileToneCurve = 50940
    EXIFIFDProfileEmbedPolicy = 50941
    EXIFIFDProfileCopyright = 50942
    EXIFIFDForwardMatrix1 = 50964
    EXIFIFDForwardMatrix2 = 50965
    EXIFIFDPreviewApplicationName = 50966
    EXIFIFDPreviewApplicationVersion = 50967
    EXIFIFDPreviewSettingsName = 50968
    EXIFIFDPreviewSettingsDigest = 50969
    EXIFIFDPreviewColorSpace = 50970
    EXIFIFDPreviewDateTime = 50971
    EXIFIFDRawImageDigest = 50972
    EXIFIFDOriginalRawFileDigest = 50973
    EXIFIFDSubTileBlockSize = 50974
    EXIFIFDRowInterleaveFactor = 50975
    EXIFIFDProfileLookTableDims = 50981
    EXIFIFDProfileLookTableData = 50982
    EXIFIFDOpcodeList1 = 51008
    EXIFIFDOpcodeList2 = 51009
    EXIFIFDOpcodeList3 = 51022
    EXIFMISCNoiseProfile = 51041
    EXIFMISCTimeCodes = 51043
    EXIFMISCFrameRate = 51044
    EXIFMISCTStop = 51058
    EXIFMISCReelName = 51081
    EXIFMISCOriginalDefaultFinalSize = 51089
    EXIFMISCOriginalBestQualitySize = 51090
    EXIFMISCOriginalDefaultCropSize = 51091
    EXIFMISCCameraLabel = 51105
    EXIFMISCProfileHueSatMapEncoding = 51107
    EXIFMISCProfileLookTableEncoding = 51108
    EXIFMISCBaselineExposureOffset = 51109
    EXIFMISCDefaultBlackRender = 51110
    EXIFMISCNewRawImageDigest = 51111
    EXIFMISCRawToPreviewGain = 51112
    EXIFMISCCacheVersion = 51114
    EXIFMISCDefaultUserCrop = 51125
    EXIFMISCNikonNEFInfo = 51157
    EXIFMISCDepthFormat = 51177
    EXIFMISCDepthNear = 51178
    EXIFMISCDepthFar = 51179
    EXIFMISCDepthUnits = 51180
    EXIFMISCDepthMeasureType = 51181
    EXIFMISCEnhanceParams = 51182
    EXIFMISCProfileGainTableMap = 52525
    EXIFMISCSemanticName = 52526
    EXIFMISCSemanticInstanceIFD = 52528
    EXIFMISCCalibrationIlluminant3 = 52529
    EXIFMISCCameraCalibration3 = 52530
    EXIFMISCColorMatrix3 = 52531
    EXIFMISCForwardMatrix3 = 52532
    EXIFMISCIlluminantData1 = 52533
    EXIFMISCIlluminantData2 = 52534
    EXIFMISCIlluminantData3 = 52535
    EXIFMISCMaskSubArea = 52536
    EXIFMISCProfileHueSatMapData3 = 52537
    EXIFMISCReductionMatrix3 = 52538
    EXIFMISCRGBTables = 52539
    EXIFMISCpadding = 59932
    EXIFMISCOffsetSchema = 59933
    EXIFMISCOwnerName = 65000
    EXIFMISCSerialNumber = 65001
    EXIFMISCLens = 65002
    EXIFMISCKDC_IFD = 65024
    EXIFMISCRawFile = 65100
    EXIFMISCConverter = 65101
    EXIFMISCWhiteBalance = 65102
    EXIFMISCExposure = 65105
    EXIFMISCShadows = 65106
    EXIFMISCBrightness = 65107
    EXIFMISCContrast = 65108
    EXIFMISCSaturation = 65109
    EXIFMISCSharpness = 65110
    EXIFMISCSmoothness = 65111
    EXIFMISCMoireFilter = 65112
End Enum

Private Type TThis
    path as string
    kind as stdShellKind
    isVirtual as boolean
    shellExtensions as Collection
End Type
Private This As TThis

'Expose type via property for external reads
Public Property Get FileType() As stdShellKind
    FileType = This.kind
End Property

'New explicit kind property
Public Property Get Kind() As stdShellKind
    Kind = This.kind
End Property

'Returns True if the shell item is deemed to be a file
'@returns - True if the shell item is a file
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").IsFile
'```
Public Property Get IsFile() As Boolean
    IsFile = This.kind = EShellFileTypeFile
End Property

'Returns True if the shell item is deemed to be a folder
'@returns - True if the shell item is a folder
'@example ```vb
'stdShell.Create("C:\my\files").IsFolder
'```
Public Property Get IsFolder() As Boolean
    IsFolder = This.kind = EShellFileTypeFolder
End Property

'Is the shell item virtual? Virtual shell items are not backed by a physical file or folder.
'@returns - True if the shell item is virtual
'@example ```vb
'stdShell.Create("C:\my\files").IsVirtual
'```
Public Property Get IsVirtual() As Boolean
    IsVirtual = This.isVirtual
End Property

'Create a new stdShell object. Kind is determined from the path, else use `AsFile` or `AsFolder` to define it.
'@constructor
'@param path - The path to wrap a shell object around.
Public Function Create(ByVal path As String) As stdShell
    Dim kind As stdShellKind
    Dim isVirtual As Boolean
    kind = privGetSystemType(path)
    isVirtual = (kind = EShellFileTypeUnknown)
    
    Set Create = New stdShell
    Call Create.protInit(path, kind, isVirtual)
End Function

'Create a virtual file handle (no I/O)
'@constructor
'@param path - Path to file
'@returns {stdShell} Virtual shell object of file
Public Function CreateFile(ByVal path As String) As stdShell
    Set CreateFile = New stdShell
    Call CreateFile.protInit(path, EShellFileTypeFile, privGetSystemType(path) = EShellFileTypeUnknown)
End Function

'Create a virtual directory handle (no I/O)
'@constructor
'@param Path to folder
'@returns {stdShell} Virtual shell object of directory
Public Function CreateFolder(ByVal path As String) As stdShell
    Set CreateFolder = New stdShell
    Call CreateFolder.protInit(path, EShellFileTypeFolder, privGetSystemType(path) = EShellFileTypeUnknown)
End Function

'Create a root shell object
'The root shell object is a virtual shell object that represents the root of the file system. It 
'contains all drives, devices and network shares.
'@constructor
'@returns - Root shell object
Public Function CreateRoot() as stdShell
    Set CreateRoot = New stdShell
    #if Mac then
        Call CreateRoot.protInit("/", EShellFileTypeFolder, True)
    #else
        Call CreateRoot.protInit("", EShellFileTypeRoot, True)
    #end if
End Function


'Create a temporary shell object. Use `AsFile` or `AsFolder` to define its type and `Ensure` to create it.
'@constructor
'@returns - Temporary shell object
Public Function CreateTemporaryItem() as stdShell
    Dim tempFolder as string: tempFolder = Environ("TEMP")
    Dim guid as string: guid = getGUID()
    Set CreateTemporaryFile = New stdShell
    Call CreateTemporaryFile.protInit(tempFolder & "\" & guid, EShellFileTypeUnknown, True)
End Function

'Initialise object
'@protected
'@param path - Path to file or folder
'@param pType - Type of shell item
'@param isVirtual - Whether the shell item is virtual
Friend Sub protInit(ByVal path As String, pType As stdShellKind, ByVal isVirtual As Boolean)
    This.path = path
    This.kind = pType
    This.isVirtual = isVirtual
    set This.shellExtensions = protGetExtensions()
End Sub

'Add a shell extension
'@static
'@param extension - Shell extension to add
Public Sub BindExtension(ByVal extension as stdIShellExtension)
    if This.shellExtensions is nothing Then Set This.shellExtensions = new Collection
    Call This.shellExtensions.Add(extension)
End Sub

'Get the shell extensions
'@protected
'@returns - Shell extensions
Public Function protGetExtensions() as Collection
    Set protGetExtensions = This.shellExtensions
End Function

'Specify that the shell item is a file. Useful when ensuring a file exists.
'@returns - File shell object
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").AsFile().Ensure()
'```
Public Function AsFile() as stdShell
    Set AsFile = CreateFile(This.path)
End Function

'Specify that the shell item is a folder. Useful when ensuring a folder exists.
'@returns - Folder shell object
'@example ```vb
'stdShell.Create("C:\my\files.stuff").AsFolder().Ensure()
'```
Public Function AsFolder() as stdShell
    Set AsFolder = CreateFolder(This.path)
End Function

'Sanitize the current path into a single legal item under the root
'Replaces illegal characters and any slashes with underscores, preserves extension for files
'@returns - New stdShell pointing to sanitized item path
Public Function ToLegalItem(optional ByVal replacement as string = "_") As stdShell
    Dim original As String: original = This.path
    original = Replace(original, "!", replacement)
    original = Replace(original, "*", replacement)
    original = Replace(original, "|", replacement)
    original = Replace(original, "<", replacement)
    original = Replace(original, ">", replacement)
    original = Replace(original, ":", replacement)
    original = Replace(original, "?", replacement)
    
    Set ToLegalItem = New stdShell
    Call ToLegalItem.protInit(original, This.kind, privGetSystemType(newPath) = EShellFileTypeUnknown)
End Function

#if not Mac then
    'Convert the shell item to a FileSystemObject
    '@returns - FileSystemObject
    '@example ```vb
    'stdShell.Create("C:\my\files\song.mp3").ToFSO()
    '```
    Public Function ToFSO() as Object
        select case This.kind
            case EShellFileTypeFile
                Set ToFSO = CreateObject("Scripting.FileSystemObject").GetFile(This.path)
            case EShellFileTypeFolder
                Set ToFSO = CreateObject("Scripting.FileSystemObject").GetFolder(This.path)
            case else
                Err.Raise 1, "", "Cannot convert this item to a FileSystemObject"
        end select
    End Function

    'Convert the shell item to an IStream
    '@returns - IStream
    '@example ```vb
    'stdShell.Create("C:\my\files\song.mp3").ToIStream()
    '```
    Public Function ToIStream() as Object
        Set ToIStream = CreateObject("ADODB.Stream")
        With ToIStream
            .Type = 1 'adTypeBinary
            .Mode = 3 'adModeReadWrite
            .Open
            .LoadFromFile This.path
            .Position = 0
        End With
    End Function

    'Create a shortcut to the shell item
    '@param path - Path to save the shortcut to
    '@returns - Success flag
    '@example ```vb
    'stdShell.Create("C:\my\files\song.mp3").SaveAsShortcut("C:\my\files\song.lnk")
    '```
    Public Function SaveAsShortcut(ByVal path as string) as Boolean
        On Error GoTo EH
        Dim shortcut As Object: Set shortcut = CreateObject("WScript.Shell").CreateShortcut(path)
        shortcut.TargetPath = This.path
        shortcut.Save

        SaveAsShortcut = True
        Exit Function
    EH:
        SaveAsShortcut = False
    End Function
#end if

'Returns the full path of the shell item
'@returns - Full path of the shell item
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").FullPath
'```
Public Property Get FullPath() as string
    FullPath = This.path
End Property

'Returns the name of the shell item
'@returns - Name of the shell item
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Name
'```
Public Property Get Name(Optional ByVal includingExtension as boolean = false) as string
    If This.kind = EShellFileTypeRoot Then Exit Property
    Name = Mid(This.path, InStrRev(This.path, "\") + 1)
    if not includingExtension then Name = Left(Name, InStrRev(Name, ".") - 1)
End Property

'Returns the extension of the shell item
'@returns - Extension of the shell item
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Extension
'```
Public Property Get Extension() as string
    If This.kind <> EShellFileTypeFile Then Exit Property
    Extension = Mid(This.path, InStrRev(This.path, ".") + 1)
End Property

'Returns the date/time the item was last modified
'@returns - Date the item was last modified
Public Property Get DateModified() as Date
    If This.kind = EShellFileTypeRoot Then Exit Property
    On Error GoTo EH
    If Exists Then
        Select Case This.kind
            Case EShellFileTypeFile
                DateModified = ToFSO().DateLastModified
            Case EShellFileTypeFolder
                DateModified = ToFSO().DateLastModified
        End Select
    End If
    Exit Property
EH:
End Property

'Returns the date/time the item was last accessed
'@returns - Date the item was last accessed
Public Property Get DateAccessed() as Date
    If This.kind = EShellFileTypeRoot Then Exit Property
    On Error GoTo EH
    If Exists Then
        Select Case This.kind
            Case EShellFileTypeFile
                DateAccessed = ToFSO().DateLastAccessed
            Case EShellFileTypeFolder
                DateAccessed = ToFSO().DateLastAccessed
        End Select
    End If
    Exit Property
EH:
End Property

'Returns the date/time the item was created
'@returns - Date the item was created
Public Property Get DateCreated() as Date
    If This.kind = EShellFileTypeRoot Then Exit Property
    On Error GoTo EH
    If Exists Then
        Select Case This.kind
            Case EShellFileTypeFile
                DateCreated = ToFSO().DateCreated
            Case EShellFileTypeFolder
                DateCreated = ToFSO().DateCreated
        End Select
    End If
    Exit Property
EH:
End Property

'Indicates whether the item has the read-only attribute
'@returns - True if read-only, False otherwise
Public Property Get IsReadOnly() as boolean
    IsReadOnly = GenericAttribute(vbReadOnly)
End Property
Public Property Let IsReadOnly(b as boolean)
    GenericAttribute(vbReadOnly) = b
End Property

'Indicates whether the item has the hidden attribute
'@returns - True if hidden, False otherwise
Public Property Get IsHidden() as boolean
    IsHidden = GenericAttribute(vbHidden)
EH:
End Property
Public Property Let IsHidden(b as boolean)
    GenericAttribute(vbHidden) = b
End Property

'Indicates whether the item has the system attribute
'@returns - True if system, False otherwise
Public Property Get IsSystemFile() as boolean
    IsSystemFile = GenericAttribute(vbSystem)
End Property
Public Property Let IsSystemFile(b as boolean)
    GenericAttribute(vbSystem) = b
End Property

'Indicates whether the item has the archive attribute / if it's ready for archiving
'@returns - True if archive, False otherwise
Public Property Get IsArchive() as boolean
    IsArchive = GenericAttribute(vbArchive)
End Property
Public Property Let IsArchive(b as boolean)
    GenericAttribute(vbArchive) = b
End Property

'Obtain the value of a property from the extended properties
'@param str - The property name to obtain
'@return Variant<Variant|Array<Variant>> - The value of the property
Public Property Get PropertyFromString(ByVal str as string) as Variant
  If Exists then
    PropertyFromString = CreateObject("Shell.Application").Namespace(0).ParseName(This.path).ExtendedProperty(str)
  End If
End Property

'Obtain the value of a property from the extended properties
'@param iProp - The property to obtain
'@return Variant<Variant|Array<Variant>> - The value of the property
Public Property Get PropertyFromEnum(ByVal iProp as EShellProperties) as Variant
  PropertyFromEnum = PropertyFromString(stdShell.protEnumToPropName(iProp))
End Property

'Obtain the value of an EXIF data property.
'@param iProp - The property to obtain
'@return - The value of the property
'@remark - It's advisable to copy the file before using the `Let` property
'@source https://github.com/KallunWillock/vbaWIA/blob/main/Properties/modWIA_EXIFMetadata.bas
Public Property Get ExifProperty(ByVal iProp as EShellExifProperties) as String
  Dim Image: Set Image = CreateObject("WIA.ImageFile")
  Call Image.LoadFile(This.path)
  set property = Image.Properties.Item(iProp) 'TODO: test, might need to search for propertyID
  If TypeName(ImageProperty.Value) = "String" Then
    ExifProperty = ImageProperty.Value
  Else
    ExifProperty = Replace(StrConv(ImageProperty.Value.BinaryData, vbUnicode), Chr(0), "")
  End If
End Property
Public Property Let ExifProperty(ByVal iProp as EShellExifProperties, ByVal v as String)
    Dim Image: Set Image = CreateObject("WIA.ImageFile")
    Call Image.LoadFile(This.path)
    set property = Image.Properties.Item(iProp) 'TODO: test, might need to search for propertyID
    if not property is nothing then
        property.Value = v
    else
        'Add new property
        Set ImageProcess = CreateObject("WIA.ImageProcess")
        Set ImageVector = CreateObject("WIA.Vector")
        ImageProcess.Filters.Add ImageProcess.FilterInfos("Exif").FilterID
        ImageProcess.Filters(1).Properties("ID") = PropertyName
        Select Case iProp
            case EXIFImageDateTimeOriginal
                Dim StringValue As String
                StringValue = Format(PropertyValue, "YYYY:MM:DD HH:MM:SS")
                ImageProcess.Filters(1).Properties("Type") = StringImagePropertyType
                ImageProcess.Filters(1).Properties("Value") = StringValue
            case else
                ImageProcess.Filters(1).Properties("Type") = VectorOfBytesImagePropertyType
                ImageVector.SetFromString PropertyValue
                ImageProcess.Filters(1).Properties("Value") = ImageVector
        end select
        Call ImageProcess.Apply(Image)
        Call Delete
        Call Image.SaveFile(This.path)
        This.isVirtual = false
    end if
End Property


' Returns True if path exists. If true, clears virtual and updates kind if unknown
'@returns - True if path exists
'@example ```vb
'If stdShell.Create("C:\my\files\song.mp3").Exists() Then
'    Debug.Print "File exists"
'End If
'```
Public Property Get Exists() As Boolean
    If This.kind = EShellFileTypeRoot Then 
        Exists = True
        Exit Property
    End if

    Dim s As String: s = This.path
    Dim k As stdShellKind: k = privGetSystemType(s)
    If k <> EShellFileTypeUnknown Then
        Exists = True
        If This.kind = EShellFileTypeUnknown Then This.kind = k
        This.isVirtual = False
    Else
        Exists = False
    End If
End Property

' Ensure file/folder exists based on intended kind; optionally create parents
'@param createParents - Whether to create parents if they don't exist
'@returns - True if the file/folder exists, false otherwise
'@example ```vb
'If stdShell.Create("C:\my\files\song.mp3").Ensure() then
'    Debug.Print "File now guaranteed to exist"
'End If
'```
Public Function Ensure(Optional ByVal createParents As Boolean = True) As Boolean
    'Root is always guaranteed to exist
    If This.kind = EShellFileTypeRoot Then 
        Ensure = True
        Exit Function
    End if
    
    'If it already exists, then we're done
    Dim s As String: s = This.path
    If Exists() Then Ensure = True: Exit Function
    Select Case This.kind
        Case EShellFileTypeFile
            If createParents Then
                Dim p As stdShell
                Set p = Me.Parent
                If Not p Is Nothing Then
                    If Not p.Ensure(True) Then Exit Function
                End If
            End If
            On Error GoTo EH
            Dim ff As Long: ff = FreeFile
            Open s For Output As #ff: Close #ff
            This.isVirtual = False
            Ensure = True
            Exit Function
        Case EShellFileTypeFolder
            On Error GoTo EH
            If createParents Then
                Dim pp As stdShell
                Set pp = Me.Parent
                If Not pp Is Nothing Then
                    If Not pp.Ensure(True) Then Exit Function
                End If
            End If
            MkDir s
            This.isVirtual = False
            Ensure = True
            Exit Function
        Case Else
            Ensure = False
            Exit Function
    End Select
EH:
    Ensure = False
End Function

' Ensure empty file; update timestamp if exists
'@returns - True if the file exists / was created, false otherwise
Public Function Touch() As Boolean
    If This.kind <> EShellFileTypeFile Then Exit Function
    Dim s As String: s = This.path
    On Error GoTo EH
    If Not Exists() Then
        Dim ff As Long: ff = FreeFile
        Open s For Output As #ff: Close #ff
        This.isVirtual = False
        Touch = True
        Exit Function
    End If
    'TODO: Refactor to use native methods
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(s) Then
        fso.GetFile(s).DateLastModified = Now
        This.isVirtual = False
        Touch = True
        Exit Function
    End If
EH:
    Touch = False
End Function

' Parent shell item
'@returns - Parent shell item
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Parent
'```
Public Property Get Parent() As stdShell
    If This.kind = EShellFileTypeRoot Then Exit Property
    Dim p As String: p = privGetParentPath(This.path)
    If p <> "" Then Set Parent = stdShell.CreateFolder(p)
End Property


' Returns the size of the file
'@returns - Size of the file
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Size
'```
Public Property Get Size() As Long
  Select Case This.kind
    Case EShellFileTypeFile
      Size = FileLen(This.path)
    Case Else
      Err.Raise 1, "", "Not implemented, use `GetSize` function instead."
  End Select
End Property

'Get the size of the file/folder
'@returns - Size of the file/folder
'@remark This is a recursive function that will return the size of the file/folder and all its children.
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").GetSize()
'```
Public Function GetSize() As Long
    Select Case This.kind
        Case EShellFileTypeFile
            GetSize = FileLen(This.path)
        Case else
            if not This.isVirtual Then
                Dim child as stdShell
                For each child in Children
                    GetSize = GetSize + child.GetSize()
                Next
                DoEvents
            end if
    End Select
End Function

'Obtains a collection of children (ignoring shellExtensions)
'@returns {Collection<stdShell>}
Public Property Get ChildrenRaw() As Collection
    Set ChildrenRaw = New Collection
    select case This.kind
        case EShellFileTypeFile
            'No children, Do nothing
        case EShellFileTypeFolder
            'Get child paths
            Dim tpath As String
            tpath = VBA.Dir(This.path & IIf(Right(This.path, 1) = "\", "*", "\*"), vbNormal Or vbDirectory Or vbHidden Or vbSystem)
            While tpath <> ""
                If tpath <> "." And tpath <> ".." Then Call ChildrenRaw.Add(stdShell.Create(This.path & "\" & tpath))
                tpath = VBA.Dir()
            Wend
            DoEvents
        case EShellFileTypeRoot
            'Use FSO for now... TODO: Use win32 API instead.
            Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
            Dim drive As Object
            For each drive in fso.Drives
                if drive.DriveLetter <> "" then
                    ChildrenRaw.Add stdShell.Create(drive.DriveLetter & ":\")
                end if
            Next
    end select
End Property

'Obtains a collection of descendents (ignoring shellExtensions)
'@returns {Collection<stdShell>}
Public Property Get DescendentsRaw() as Collection
    Dim ret As Collection: set ret = new Collection
    Dim child As stdShell
    For each child in ChildrenRaw
        ret.add child
        For each descendent in child.DescendentsRaw
            ret.add descendent
        Next
    Next
    Set DescendentsRaw = ret
End Property


'Obtains a collection of children (including shellExtensions)
'@returns Collection<stdShell | stdIShellExtension> - Collection of children
Public Property Get Children() as Collection
    Dim raw As Collection: Set raw = Me.ChildrenRaw
    Dim consumed as Object: set consumed = CreateObject("Scripting.Dictionary")
    Dim ret as Collection: set ret = new Collection
    
    'Extend children
    Dim ext as stdIShellExtension
    For each ext in This.shellExtensions
        Dim child As stdShell
        For each child in raw
            If not consumed.exists(child.FullPath) Then 
                If ext.IsExtensionOf(child) Then
                    Dim extendedChild as stdIShellExtension
                    set extendedChild = ext.CreateExtension(child)
                    ret.add extendedChild
                    consumed(child.FullPath) = true
                    Dim desc as Variant 'stdIShellExtension or stdShell
                    For each desc in extendedChild.DescendentsRaw
                        consumed(desc.FullPath) = true
                    next
                end if
            end if
        Next
    Next

    'Add remaining children that were not extended
    For each child in raw
        If not consumed(child.FullPath) Then
            ret.add child
        end if
    next
    
    Set Children = ret
End Property

'Attempts to read data from a file
'@returns {Byte()} - Bytes read from file
Public Function Read() As Byte()
    If This.kind = EShellFileTypeFile Then
        Dim ff As Long: ff = FreeFile
        Open This.path For Input As #ff
            Read = Input(LOF(ff), #ff)
        Close #ff
    Else
        Err.Raise 1, "", "Cannot read this item as it is not a file"
    End If
End Function

'Attempts to read data from a file
'@returns {Byte()} - Bytes read from file
Public Function ReadText() As String
    If This.kind = EShellFileTypeFile Then
        Dim ff As Long: ff = FreeFile
        Open This.path For Input As #ff
            ReadText = Input$(LOF(ff), #ff)
        Close #ff
    Else
        Err.Raise 1, "", "Cannot read this item as it is not a file"
    End If
End Function

'Attempts to append data onto the end of a file
'@param data {ByVal Variant} - Data to append to file
'@returns {Boolean} - Success flag
Public Function Append(ByVal data As Variant) As Boolean
    On Error GoTo EH
        If This.kind = EShellFileTypeFile Then
            Dim ff As Long: ff = FreeFile
            Open This.path For Append As #ff
                Seek #ff, LOF(ff) + 1
                Put #ff, , CStr(data)
            Close #ff
        Else
            Err.Raise 1, "", "Cannot append data to this item as it is not a file"
        End If
        Exit Function
EH:
    Append = False
End Function

'Appends a line to the file
'@param line {ByVal Variant} - Line to append to file
'@returns {Boolean} - Success flag
Public Function AppendLine(ByVal line As string) As Boolean
    If This.kind = EShellFileTypeFile Then
        AppendLine = Append(data & vbCrLf)
    Else
        Err.Raise 1, "", "Cannot append line to this item as it is not a file"
    End If
End Function

'Attempts to overwrite the contents of the file
'@param data {ByVal Variant} - Data to write to file
'@returns {Boolean} - Success flag
Public Function Overwrite(ByVal data As Variant) As Boolean
    On Error GoTo EH
        If This.kind = EShellFileTypeFile Then
            Dim ff As Long: ff = FreeFile
            Open This.path For Binary As #ff
                Put #ff, , data
            Close #ff
        Else
            Err.Raise 1, "", "Cannot write data to this item as it is not a file"
        End If
        Exit Function
EH:
    Overwrite = False
End Function

'Attempts to clear the file
'@returns {Boolean} - Success flag
Public Function Clear() As Boolean
    On Error GoTo EH
        If This.kind = EShellFileTypeFile Then
            Dim ff As Long: ff = FreeFile
            Open This.path For Output As #ff
            Close #ff
        Else
            Err.Raise 1, "", "Cannot clear data from this item as it is not a file"
        End If
        Exit Function
EH:
    Clear = False
End Function

'Obtain a hash value for file contents
'@param {stdShellHashType} The type of hash to create
'@returns Hex string representing hash.
'@remark {1} MD5 and SHA1 kept for legacy support.
Public Function Hash(ByVal iHashType As stdShellHashType) As String
    'Obtain hash encoder
    Dim hashEncoder As Object, iHashLength As Long
    Select Case iHashType
        Case HashTypeMD5
            Set hashEncoder = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
            iHashLength = 16
        Case HashTypeSHA1
            Set hashEncoder = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
            iHashLength = 20
        Case HashTypeSHA256
            Set hashEncoder = CreateObject("System.Security.Cryptography.SHA256Managed")
            iHashLength = 32
    End Select
    
    Dim blockSize As Long: blockSize = 2 ^ 16
    Dim ff As Long: ff = FreeFile
    
    'Progressively build hash 1024 bytes at a time
    Open This.path For Binary Access Read As #ff
        If LOF(ff) < blockSize Then blockSize = ((LOF(ff) + 1024) \ 1024) * 1024
        Dim buffer() As Byte: ReDim buffer(blockSize - 1)
        
        Dim i As Long
        For i = 1 To LOF(ff) \ blockSize
            Get ff, , buffer
            hashEncoder.TransformBlock buffer, 0, blockSize, buffer, 0
        Next
        Get ff, , buffer
        hashEncoder.TransformFinalBlock buffer, 0, LOF(ff) Mod blockSize
        buffer = hashEncoder.Hash
        hashEncoder.Clear
    Close #ff
    
    'Output hash as hex string
    Hash = ""
    For i = 0 To (iHashLength - 1)
        Hash = Hash & IIf(buffer(i) < 16, "0", "") & Hex(buffer(i))
    Next
End Function

'Finds the first shell item that matches the callable
'@param callable - Callable to match shell item against
'@param bRecurse - Whether to recurse into children
'@returns - First shell item that matches the callable
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").Find(Callable)
'```
Public Function Find(ByVal callable As stdICallable, Optional ByVal bRecurse As Boolean = True) As stdShell
    If This.kind = EShellFileTypeFolder Then
        'Get children as collection
        Dim col As Collection: Set col = Children
        
        Dim shellItem As stdShell
        While Not CollectionShift(col, shellItem) Is Nothing
            Dim child As stdShell
            For Each child In shellItem.Children
                If callable.Run(child) Then
                    Set Find = child
                    Exit Function
                End If
                
                If child.Kind = EShellFileTypeFolder Then col.add child
            Next
        Wend
    End If
End Function

'Finds all shell items that match the callable
'@param callable - Callable to match shell item against
'@returns - Collection of shell items that match the callable
'@example ```vb
'stdShell.Create("C:\my\files\song.mp3").FindAll(Callable)
'```
Public Function FindAll(ByVal callable As stdICallable) As Collection
    Dim vRet As Collection: Set vRet = New Collection
    If This.kind = EShellFileTypeFolder Then
        'Get children as collection
        Dim col As Collection: Set col = Children
        
        Dim shellItem As stdShell
        While Not CollectionShift(col, shellItem) Is Nothing
            Dim child As stdShell
            For Each child In shellItem.Children
                If callable.Run(child) Then
                    vRet.Add child
                End If
                
                If child.Kind = EShellFileTypeFolder Then col.add child
            Next
        Wend
    End If
    Set FindAll = vRet
End Function

'Deletes the file/folder
'@returns {Boolean} - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").Delete()
'```
Public Function Delete() As Boolean
    On Error GoTo EH
        If Exists Then
            select case This.kind
                case EShellFileTypeRoot
                    Exit Function
                case EShellFileTypeFolder
                    Dim child as stdShell
                    For each child in Children
                        child.Delete
                    next
                    RmDir This.path
                case EShellFileTypeFile
                    Kill This.path
                case else
                    Err.Raise 1, "", "Cannot delete this item as it is not a file"
            end select
            This.isVirtual = True
        End If
        Delete = True
        Exit Function
EH:
    Delete = False
End Function

'Copies the file/folder to a new path
'@param path - Path to copy the file/folder to
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").CopyTo("C:\my\files\song2.mp3")
'```
Public Function CopyTo(ByVal path As String) As Boolean
    On Error GoTo EH
        If Exists Then
            select case This.kind
                case EShellFileTypeFile
                    Call FileCopy(This.path, path)
                case EShellFileTypeFolder
                    With CreateObject("Scripting.FileSystemObject")
                        Call .CopyFolder(This.path, path)
                    End With
                case else
                    Err.Raise 1, "", "Cannot copy this item as it is not a file or folder"
            end select
            CopyTo = True
        End If
        Exit Function
EH:
    CopyTo = False
End Function

'Copies the file/folder to a new folder
'@param folderPath - Path to copy the file/folder to
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").CopyToFolder("C:\my\files")
'```
Public Function CopyToFolder(ByVal folderPath As String) As Boolean
    if right(folderPath, 1) <> "\" then folderPath = folderPath & "\"
    Call CopyTo(folderPath & Name(true))
End Function


'Moves the file/folder to a new path
'@param path - Path to move the file/folder to
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").MoveTo("C:\my\files\song2.mp3")
'```
Public Function MoveTo(ByVal path As String) As Boolean
    On Error GoTo EH
        If Exists Then
            Call CopyTo(path)
            Call Delete()
            MoveTo = True
        End If
        Exit Function
EH:
    MoveTo = False
End Function


'Moves the file/folder to a new folder
'@param folderPath - Path to move the file/folder to
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").MoveToFolder("C:\my\files")
'```
Public Function MoveToFolder(ByVal folderPath As String) As Boolean
    if right(folderPath, 1) <> "\" then folderPath = folderPath & "\"
    Call MoveTo(folderPath & Name(true))
End Function


'Renames the file/folder
'@param newName - New name of the file/folder
'@returns - Success flag
'@example ```vb
'Call stdShell.Create("C:\my\files\song.mp3").Rename("song2.mp3")
'```
Public Function Rename(ByVal newName As String) As Boolean
    On Error GoTo EH
        If Exists Then
            Call MoveTo(Parent.FullPath & "\" & newName & Extension)
            Rename = True
        End If
        Exit Function
EH:
    Rename = False
End Function

'====================
' Private helpers
'====================
'Shift a value out of a collection
'@param col - Collection to shift value out of
'@param shellItem - Shell item shifted out of collection
'@returns - Shell item shifted out of collection
Private Function CollectionShift(ByRef col As Collection, ByRef shellItem As stdShell) As stdShell
    Set CollectionShift = col.Item(1)
    Set shellItem = col.Item(1)
    Call col.Remove(1)
End Function

'Get the parent path of a given path
'@param s - Path to get the parent path of
'@returns - Parent path of the given path
Private Function privGetParentPath(ByVal s As String) As String
    Dim path As String: path = s
    If VBA.Len(path) = 0 Then Exit Function
    If Right$(path, 1) = "\" Then
        If Not privIsDriveOrShare(path) Then path = Left$(path, Len(path) - 1)
    End If
    Dim i As Long: i = InStrRev(path, "\")
    If i <= 0 Then Exit Function
    Dim parent As String: parent = Left$(path, i - 1)
    If parent = "" Then Exit Function
    If Mid$(path, 2, 1) = ":" And i = 3 Then parent = Left$(path, 3)
    If privIsDriveOrShare(parent) Then
        ' Root has no parent; return empty to stop recursion
        privGetParentPath = ""
        Exit Function
    End If
    privGetParentPath = parent
End Function

'Check if a path is a drive or share
'@param s - Path to check if it is a drive or share
'@returns - True if the path is a drive or share, false otherwise
Private Function privIsDriveOrShare(ByVal s As String) As Boolean
    If Len(s) = 3 And Mid$(s, 2, 1) = ":" And Right$(s, 1) = "\" Then
        privIsDriveOrShare = True
    ElseIf Left$(s, 2) = "\\" Then
        Dim parts() As String: parts = Split(Mid$(s, 3), "\")
        privIsDriveOrShare = (UBound(parts) < 1)
    Else
        privIsDriveOrShare = False
    End If
End Function

'Get the system type of a given path
'@param path - Path to get the system type of
'@returns - System type of the given path
Private Function privGetSystemType(ByVal path As String) As stdShellKind
    On Error GoTo EH
    If VBA.LenB(VBA.Dir(path, vbDirectory Or vbNormal Or vbHidden Or vbSystem)) = 0 Then
        privGetSystemType = EShellFileTypeUnknown
        Exit Function
    End If
    Dim attr As VbFileAttribute: attr = VBA.GetAttr(path)
    If (attr And VbFileAttribute.vbDirectory) = VbFileAttribute.vbDirectory Or _
       (attr And VbFileAttribute.vbVolume) = VbFileAttribute.vbVolume Then
        privGetSystemType = EShellFileTypeFolder
    Else
        privGetSystemType = EShellFileTypeFile
    End If
    Exit Function
EH:
    privGetSystemType = EShellFileTypeUnknown
End Function

'Generates a new GUID
'@returns - A new GUID
Private Function getGUID() As String
  Call Randomize 'Ensure random GUID generated
  getGUID = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx"
  getGUID = Replace(getGUID, "y", Hex(Rnd() And &H3 Or &H8))
  Dim i As Long: For i = 1 To 30
    getGUID = Replace(getGUID, "x", Hex$(Int(Rnd() * 16)), 1, 1)
  Next
End Function

'Indicates whether the item has the read-only attribute
'@returns - True if read-only, False otherwise
Private Property Get GenericAttribute(Byval iAttr as vbFileAttribute) as boolean
    If This.kind = EShellFileTypeRoot Then Exit Property
    On Error GoTo EH
    If Exists Then
        Dim attr As VbFileAttribute: attr = VBA.GetAttr(This.path)
        GenericAttribute = (attr And iAttr) = iAttr
    End If
EH:
End Property
Private Property Let GenericAttribute(Byval iAttr as vbFileAttribute, b as boolean)
    If This.kind = EShellFileTypeRoot Then Exit Property
    On Error GoTo EH
    If Exists Then
        Dim attr As VbFileAttribute: attr = VBA.GetAttr(This.path)
        if b then
            attr = attr Or iAttr
        else
            attr = attr And Not iAttr
        end if
    End If
EH:
End Property

'Obtain the name of a property from the extended properties
'@protected
'@param prop - The property to obtain
'@return String - The name of the property as utilised in Shell.Application
Public Function protEnumToPropName(ByVal prop as stdShellExtendedProps) As String
    static map as variant
    if isEmpty(map) Then
        Dim s as string
        'System/Generic
        s = s & "$AcquisitionID,$ApplicationDefinedProperties,$ApplicationName,$AppZoneIdentifier,$Author,$CachedFileUpdaterContentIdForConflictResolution"
        s = s & ",$CachedFileUpdaterContentIdForStream,$Capacity,$Category,$Comment,$Company,$ComputerName,$ContainedItems,$ContentStatus,$ContentType,$Copyright"
        s = s & ",$CreatorAppId,$CreatorOpenWithUIOptions,$DataObjectFormat,$DateAccessed,$DateAcquired,$DateArchived,$DateCompleted,$DateCreated,$DateImported"
        s = s & ",$DateModified,$DefaultSaveLocationDisplay,$DueDate,$EndDate,$ExpandoProperties,$FileAllocationSize,$FileAttributes,$FileCount,$FileDescription"
        s = s & ",$FileExtension,$FileFRN,$FileName,$FileOfflineAvailabilityStatus,$FileOwner,$FilePlaceholderStatus,$FileVersion,$FindData,$FlagColor,$FlagColorText"
        s = s & ",$FlagStatus,$FlagStatusText,$FolderKind,$FolderNameDisplay,$FreeSpace,$FullText,$HighKeywords,$ImageParsingName,$Importance,$ImportanceText,$IsAttachment"
        s = s & ",$IsDefaultNonOwnerSaveLocation,$IsDefaultSaveLocation,$IsDeleted,$IsEncrypted,$IsFlagged,$IsFlaggedComplete,$IsIncomplete,$IsLocationSupported"
        s = s & ",$IsPinnedToNameSpaceTree,$IsRead,$IsSearchOnlyItem,$IsSendToTarget,$IsShared,$ItemAuthors,$ItemClassType,$ItemDate,$ItemFolderNameDisplay"
        s = s & ",$ItemFolderPathDisplay,$ItemFolderPathDisplayNarrow,$ItemName,$ItemNameDisplay,$ItemNameDisplayWithoutExtension,$ItemNamePrefix"
        s = s & ",$ItemNameSortOverride,$ItemParticipants,$ItemPathDisplay,$ItemPathDisplayNarrow,$ItemSubType,$ItemType,$ItemTypeText,$ItemUrl,$Keywords,$Kind,$KindText"
        s = s & ",$Language,$LastSyncError,$LastWriterPackageFamilyName,$LowKeywords,$MediumKeywords,$MileageInformation,$MIMEType,$Null,$OfflineAvailability,$OfflineStatus"
        s = s & ",$OriginalFileName,$OwnerSID,$ParentalRating,$ParentalRatingReason,$ParentalRatingsOrganization,$ParsingBindContext,$ParsingName,$ParsingPath,$PerceivedType"
        s = s & ",$PercentFull,$Priority,$PriorityText,$Project,$ProviderItemID,$Rating,$RatingText,$RemoteConflictingFile,$Sensitivity,$SensitivityText,$SFGAOFlags,$SharedWith"
        s = s & ",$ShareUserRating,$SharingStatus,$Shell,$SimpleRating,$Size,$SoftwareUsed,$SourceItem,$SourcePackageFamilyName,$StartDate,$Status"
        s = s & ",$StorageProviderCallerVersionInformation,$StorageProviderError,$StorageProviderFileChecksum,$StorageProviderFileIdentifier,$StorageProviderFileRemoteUri"
        s = s & ",$StorageProviderFileVersion,$StorageProviderFileVersionWaterline,$StorageProviderId,$StorageProviderShareStatuses,$StorageProviderSharingStatus"
        s = s & ",$StorageProviderStatus,$Subject,$SyncTransferStatus,$Thumbnail,$ThumbnailCacheId,$ThumbnailStream,$Title,$TitleSortOverride,$TotalFileSize,$Trademarks"
        s = s & ",$TransferOrder,$TransferPosition,$TransferSize,$VolumeId,$ZoneIdentifier"
        s = replace(s, "$", "System.")
        
        'Document
        s = s & ",$ByteCount,$CharacterCount,$ClientID,$Contributor,$DateCreated,$DatePrinted,$DateSaved,$Division,$DocumentID,$HiddenSlideCount,$LastAuthor,$LineCount,$Manager"
        s = s & ",$MultimediaClipCount,$NoteCount,$PageCount,$ParagraphCount,$PresentationFormat,$RevisionNumber,$Security,$SlideCount,$Template,$TotalEditingTime,$Version,$WordCount"
        s = replace(s, "$", "System.Document.")

        'Video
        s = s & ",$Compression,$Director,$EncodingBitrate,$FourCC,$FrameHeight,$FrameRate,$FrameWidth,$HorizontalAspectRatio,$IsSpherical,$IsStereo,$Orientation,$SampleSize"
        s = s & ",$StreamName,$StreamNumber,$TotalBitrate,$TranscodedForSync,$VerticalAspectRatio"
        s = replace(s, "$", "System.Video.")

        'Photo
        s = s & ",$Aperture,$ApertureDenominator,$ApertureNumerator,$Brightness,$BrightnessDenominator,$BrightnessNumerator,$CameraManufacturer,$CameraModel"
        s = s & ",$CameraSerialNumber,$Contrast,$ContrastText,$DateTaken,$DigitalZoom,$DigitalZoomDenominator,$DigitalZoomNumerator,$Event,$EXIFVersion,$ExposureBias"
        s = s & ",$ExposureBiasDenominator,$ExposureBiasNumerator,$ExposureIndex,$ExposureIndexDenominator,$ExposureIndexNumerator,$ExposureProgram,$ExposureProgramText"
        s = s & ",$ExposureTime,$ExposureTimeDenominator,$ExposureTimeNumerator,$Flash,$FlashEnergy,$FlashEnergyDenominator,$FlashEnergyNumerator,$FlashManufacturer"
        s = s & ",$FlashModel,$FlashText,$FNumber,$FNumberDenominator,$FNumberNumerator,$FocalLength,$FocalLengthDenominator,$FocalLengthInFilm,$FocalLengthNumerator"
        s = s & ",$FocalPlaneXResolution,$FocalPlaneXResolutionDenominator,$FocalPlaneXResolutionNumerator,$FocalPlaneYResolution,$FocalPlaneYResolutionDenominator"
        s = s & ",$FocalPlaneYResolutionNumerator,$GainControl,$GainControlDenominator,$GainControlNumerator,$GainControlText,$ISOSpeed,$LensManufacturer,$LensModel"
        s = s & ",$LightSource,$MakerNote,$MakerNoteOffset,$MaxAperture,$MaxApertureDenominator,$MaxApertureNumerator,$MeteringMode,$MeteringModeText,$Orientation"
        s = s & ",$OrientationText,$PeopleNames,$PhotometricInterpretation,$PhotometricInterpretationText,$ProgramMode,$ProgramModeText,$RelatedSoundFile,$Saturation"
        s = s & ",$SaturationText,$Sharpness,$SharpnessText,$ShutterSpeed,$ShutterSpeedDenominator,$ShutterSpeedNumerator,$SubjectDistance,$SubjectDistanceDenominator"
        s = s & ",$SubjectDistanceNumerator,$TagViewAggregate,$TranscodedForSync,$WhiteBalance,$WhiteBalanceText"
        s = replace(s, "$", "System.Photo.")

        'Music
        s = s & ",$AlbumArtist,$AlbumArtistSortOverride,$AlbumID,$AlbumTitle,$AlbumTitleSortOverride,$Artist,$ArtistSortOverride,$BeatsPerMinute,$Composer,$ComposerSortOverride"
        s = s & ",$Conductor,$ContentGroupDescription,$DiscNumber,$DisplayArtist,$Genre,$InitialKey,$IsCompilation,$Lyrics,$Mood,$PartOfSet,$Period,$SynchronizedLyrics,$TrackNumber"
        s = replace(s, "$", "System.Music.")

        'EMail/Message
        s = s & ",$AttachmentContents,$AttachmentNames,$BccAddress,$BccName,$CcAddress,$CcName,$ConversationID,$ConversationIndex,$DateReceived,$DateSent,$Flags,$FromAddress"
        s = s & ",$FromName,$HasAttachments,$IsFwdOrReply,$MessageClass,$Participants,$ProofInProgress,$SenderAddress,$SenderName,$Store,$ToAddress,$ToDoFlags,$ToDoTitle,$ToName"
        s = replace(s, "$", "System.Message.")

        'Media
        s = s & ",$AuthorUrl,$AverageLevel,$ClassPrimaryID,$ClassSecondaryID,$CollectionGroupID,$CollectionID,$ContentDistributor,$ContentID,$CreatorApplication"
        s = s & ",$CreatorApplicationVersion,$DateEncoded,$DateReleased,$DlnaProfileID,$Duration,$DVDID,$EncodedBy,$EncodingSettings,$EpisodeNumber,$FrameCount,$MCDI"
        s = s & ",$MetadataContentProvider,$Producer,$PromotionUrl,$ProtectionType,$ProviderRating,$ProviderStyle,$Publisher,$SeasonNumber,$SeriesName,$SubscriptionContentId"
        s = s & ",$SubTitle,$ThumbnailLargePath,$ThumbnailLargeUri,$ThumbnailSmallPath,$ThumbnailSmallUri,$UniqueFileIdentifier,$UserNoAutoInfo,$UserWebUrl,$Writer,$Year"
        s = replace(s, "$", "System.Media.")

        'Image
        s = s & ",$BitDepth,$ColorSpace,$CompressedBitsPerPixel,$CompressedBitsPerPixelDenominator,$CompressedBitsPerPixelNumerator,$Compression,$CompressionText,$Dimensions"
        s = s & ",$HorizontalResolution,$HorizontalSize,$ImageID,$ResolutionUnit,$VerticalResolution,$VerticalSize"
        s = replace(s, "$", "System.Image.")
        
        map = split(s, ",")
    end if
    protEnumToPropName = map(prop)
Exit Function