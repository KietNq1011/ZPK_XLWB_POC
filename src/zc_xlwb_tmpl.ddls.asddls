@AccessControl.authorizationCheck: #NOT_REQUIRED
@Metadata.allowExtensions: true
@Metadata.ignorePropagatedAnnotations: true
@EndUserText.label: 'XLWB Excel template'
@ObjectModel.semanticKey: [ 'FormName' ]
define root view entity ZC_XLWB_TMPL
  provider contract transactional_query
  as projection on ZR_XLWB_TMPL
{
  key FormName,
      Descr,
      Engine,

      @Semantics.mimeType: true
      MimeType,
      FileName,
      IsActive,

      @Semantics.largeObject: { mimeType: 'MimeType',
                                fileName: 'FileName',
                                acceptableMimeTypes: [ 'application/vnd.ms-excel',
                                                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                                       'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'image/png', 'text/xml' ],
                                contentDispositionPreference: #ATTACHMENT }
      Template,

      SourceClass,
      SampleJson,

      @Semantics.mimeType: true
      PreviewMime,
      PreviewName,
      @Semantics.largeObject: { mimeType: 'PreviewMime',
                                fileName: 'PreviewName',
                                contentDispositionPreference: #ATTACHMENT }
      PreviewFile,

      @Semantics.user.createdBy: true
      CreatedBy,
      @Semantics.systemDateTime.createdAt: true
      CreatedAt,
      @Semantics.user.localInstanceLastChangedBy: true
      LocalLastChangedBy,
      @Semantics.systemDateTime.localInstanceLastChangedAt: true
      LocalLastChangedAt,
      @Semantics.systemDateTime.lastChangedAt: true
      LastChangedAt
}
