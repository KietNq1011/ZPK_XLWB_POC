@AccessControl.authorizationCheck: #NOT_REQUIRED
@Metadata.allowExtensions: true
@EndUserText.label: 'XLWB Excel template (root)'
define root view entity ZR_XLWB_TMPL
  as select from zxlwb_tmpl
{
  key form_name            as FormName,
      descr                as Descr,
      engine               as Engine,

      @Semantics.mimeType: true
      mime_type            as MimeType,
      file_name            as FileName,
      is_active            as IsActive,

      @Semantics.largeObject: { mimeType: 'MimeType',
                                fileName: 'FileName',
                                acceptableMimeTypes: [ 'application/vnd.ms-excel',
                                                       'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                                       'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'image/png', 'text/xml' ],
                                contentDispositionPreference: #ATTACHMENT }
      template             as Template,

      source_class         as SourceClass,
      sample_json          as SampleJson,

      @Semantics.mimeType: true
      preview_mime         as PreviewMime,
      preview_name         as PreviewName,
      @Semantics.largeObject: { mimeType: 'PreviewMime',
                                fileName: 'PreviewName',
                                contentDispositionPreference: #ATTACHMENT }
      preview_file         as PreviewFile,

      @Semantics.user.createdBy: true
      created_by           as CreatedBy,
      @Semantics.systemDateTime.createdAt: true
      created_at           as CreatedAt,
      @Semantics.user.localInstanceLastChangedBy: true
      last_changed_by      as LocalLastChangedBy,
      @Semantics.systemDateTime.localInstanceLastChangedAt: true
      last_changed_at      as LocalLastChangedAt,
      @Semantics.systemDateTime.lastChangedAt: true
      last_changed_glo     as LastChangedAt
}
