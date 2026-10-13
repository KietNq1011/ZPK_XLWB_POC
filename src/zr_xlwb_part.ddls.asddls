@AccessControl.authorizationCheck: #NOT_REQUIRED
@Metadata.allowExtensions: true
@EndUserText.label: 'XLWB XML part từ DDIC (root)'
define root view entity ZR_XLWB_PART
  as select from zxlwb_part
{
  key part_name        as PartName,
      descr            as Descr,
      ddic_name        as DdicName,
      ddic_kind        as DdicKind,
      root_name        as RootName,
      namespace        as Namespace,
      sample_json      as SampleJson,

      @Semantics.mimeType: true
      part_mime        as PartMime,
      part_fname       as PartFileName,
      @Semantics.largeObject: { mimeType: 'PartMime',
                                fileName: 'PartFileName',
                                contentDispositionPreference: #ATTACHMENT }
      part_file        as PartFile,

      @Semantics.user.createdBy: true
      created_by       as CreatedBy,
      @Semantics.systemDateTime.createdAt: true
      created_at       as CreatedAt,
      @Semantics.user.localInstanceLastChangedBy: true
      last_changed_by  as LocalLastChangedBy,
      @Semantics.systemDateTime.localInstanceLastChangedAt: true
      last_changed_at  as LocalLastChangedAt,
      @Semantics.systemDateTime.lastChangedAt: true
      last_changed_glo as LastChangedAt
}
