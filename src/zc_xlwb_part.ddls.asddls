@AccessControl.authorizationCheck: #NOT_REQUIRED
@Metadata.ignorePropagatedAnnotations: true
@EndUserText.label: 'XLWB XML part từ DDIC'
@ObjectModel.semanticKey: [ 'PartName' ]
@UI.headerInfo: { typeName: 'XML Part',
                  typeNamePlural: 'XML Parts',
                  title: { type: #STANDARD, value: 'PartName' },
                  description: { type: #STANDARD, value: 'Descr' } }
@UI.presentationVariant: [ { sortOrder: [ { by: 'PartName', direction: #ASC } ] } ]
define root view entity ZC_XLWB_PART
  provider contract transactional_query
  as projection on ZR_XLWB_PART
{
      @UI.facet: [ { id: 'General',  purpose: #STANDARD, type: #COLLECTION,
                     label: 'XML Part', position: 10 },
                   { id: 'Head',     parentId: 'General', purpose: #STANDARD,
                     type: #FIELDGROUP_REFERENCE, label: 'General',
                     position: 10, targetQualifier: 'Head' },
                   { id: 'Gen',      parentId: 'General', purpose: #STANDARD,
                     type: #FIELDGROUP_REFERENCE, label: 'Generated output',
                     position: 20, targetQualifier: 'Gen' },
                   { id: 'Admin',    parentId: 'General', purpose: #STANDARD,
                     type: #FIELDGROUP_REFERENCE, label: 'Administrative data',
                     position: 30, targetQualifier: 'Admin' } ]

      @UI: { lineItem:       [ { position: 10, importance: #HIGH },
                               { type: #FOR_ACTION, dataAction: 'generateXml', label: 'Generate XML Part' } ],
             identification: [ { position: 10 },
                               { type: #FOR_ACTION, dataAction: 'generateXml', label: 'Generate XML Part' } ],
             fieldGroup:     [ { qualifier: 'Head', position: 10 } ],
             selectionField: [ { position: 10 } ] }
      @EndUserText.label: 'Generate XML Part'
  key PartName,

      @UI: { lineItem:       [ { position: 20, importance: #MEDIUM } ],
             identification: [ { position: 20 } ],
             fieldGroup:     [ { qualifier: 'Head', position: 20 } ] }
      @EndUserText.label: 'Description'
      Descr,

      @UI: { lineItem:       [ { position: 30, importance: #HIGH } ],
             identification: [ { position: 30 } ],
             fieldGroup:     [ { qualifier: 'Head', position: 30 } ],
             selectionField: [ { position: 20 } ] }
      @EndUserText.label: 'DDIC Name'
      DdicName,

      @UI: { lineItem:       [ { position: 40, importance: #MEDIUM } ],
             identification: [ { position: 40 } ],
             fieldGroup:     [ { qualifier: 'Head', position: 40 } ] }
      @EndUserText.label: 'DDIC Kind'
      DdicKind,

      @UI: { identification: [ { position: 50 } ],
             fieldGroup:     [ { qualifier: 'Head', position: 50 } ] }
      @EndUserText.label: 'Root Name'
      RootName,

      @UI: { identification: [ { position: 60 } ],
             fieldGroup:     [ { qualifier: 'Head', position: 60 } ] }
      @EndUserText.label: 'Name space'
      Namespace,

      @UI: { identification: [ { position: 70 } ],
             fieldGroup:     [ { qualifier: 'Gen', position: 20 } ] }
      @UI.multiLineText: true
      @EndUserText.label: 'Sample Json'
      SampleJson,

      @Semantics.mimeType: true
      @EndUserText.label: 'Part Mime'
      PartMime,

      @UI.hidden: true
      @EndUserText.label: 'Part File Name'
      PartFileName,

      @UI: { lineItem:       [ { position: 50, importance: #HIGH, label: 'XML Part file' } ],
             identification: [ { position: 80 } ],
             fieldGroup:     [ { qualifier: 'Gen', position: 10 } ] }
      @Semantics.largeObject: { mimeType: 'PartMime',
                                fileName: 'PartFileName',
                                contentDispositionPreference: #ATTACHMENT }
      @EndUserText.label: 'XML Part file'
      PartFile,

      @UI: { fieldGroup: [ { qualifier: 'Admin', position: 10 } ] }
      @Semantics.user.createdBy: true
      CreatedBy,
      @UI: { fieldGroup: [ { qualifier: 'Admin', position: 20 } ] }
      @Semantics.systemDateTime.createdAt: true
      CreatedAt,
      @UI: { fieldGroup: [ { qualifier: 'Admin', position: 30 } ] }
      @Semantics.user.localInstanceLastChangedBy: true
      LocalLastChangedBy,
      @UI: { fieldGroup: [ { qualifier: 'Admin', position: 40 } ] }
      @Semantics.systemDateTime.localInstanceLastChangedAt: true
      LocalLastChangedAt,
      @UI.hidden: true
      @Semantics.systemDateTime.lastChangedAt: true
      LastChangedAt
}
