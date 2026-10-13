@AccessControl.authorizationCheck: #NOT_REQUIRED
@Metadata.allowExtensions: true
@EndUserText.label: 'XLWB example catalog'
@ObjectModel.usageType: { serviceQuality: #X, sizeCategory: #S, dataClass: #CUSTOMIZING }
define root view entity ZC_XLWB_DEMO
  as select from zxlwb_demo
{
  key example_id       as ExampleId,
      seq              as Seq,
      xlwb_ref         as XlwbRef,
      title            as Title,
      engine_type      as EngineType,
      purpose          as Purpose,
      template_form    as TemplateForm,
      step_context     as StepContext,
      step_excel       as StepExcel,
      step_template    as StepTemplate,
      step_engine      as StepEngine,
      notes            as Notes,

      file_name        as FileName,
      @Semantics.mimeType: true
      mime_type        as MimeType,

      @Semantics.largeObject: { mimeType: 'MimeType',
                                fileName: 'FileName',
                                contentDispositionPreference: #ATTACHMENT }
      content          as Content,

      @Semantics.user.createdBy: true
      created_by       as CreatedBy,
      @Semantics.systemDateTime.createdAt: true
      created_at       as CreatedAt
}
