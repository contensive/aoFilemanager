
using System;
using Contensive.BaseClasses;
using Contensive.Models.Db;

namespace Contensive.Addons.aoFileManager2 {
    namespace Models.Db {
        public class SampleModel : DesignBlockBaseModel {
            /// <summary>
            /// table definition
            /// </summary>
            public static readonly DbBaseTableMetadataModel tableMetadata = new DbBaseTableMetadataModel("blank", "blank", "default", false);
            //====================================================================================================
            //
            // -- sample model properties. Each property name matches a database field, and the type matches the Contensive use case.
            public string imageFilename { get; set; }
            public string headline { get; set; }
            public string embed { get; set; }
            public string description { get; set; }
            public string buttonText { get; set; }
            public string buttonUrl { get; set; }
            public int imageAspectRatioId { get; set; }
            public string btnStyleSelector { get; set; }
            //
            // -- Sample property Types
            //public string sampleDbFieldText { get; set; }
            //public int sampleDbFieldInteger { get; set; }
            //public bool sampleDbFieldBoolean { get; set; }
            //public DateTime sampleDbFieldDate { get; set; }
            //public double sampleDbFieldNumber { get; set; }
            //
            // -- Sample file property Types
            //public FieldTypeCSSFile sampleDbFieldCssFile { get; set; }
            //public FieldTypeHTMLFile sampleDbFielHtmlFile { get; set; }
            //public FieldTypeJavascriptFile sampleDbFieldJavascriptFile { get; set; }
            //public FieldTypeTextFile sampleDbFieldTextFile { get; set; }
            //
            // -- Sample nullable property Types. These properties will save null if not initialized.
            //public int? sampleDbFieldNullableInteger { get; set; }
            //public bool? sampleDbFieldNullableBoolean { get; set; }
            //public DateTime? sampleDbFieldNullableDate { get; set; }
            //public double? sampleDbFieldNullableNumber { get; set; }
            // 
            // ====================================================================================================
            public static SampleModel createOrAddSettings(CPBaseClass cp, string settingsGuid) {
                SampleModel result = create<SampleModel>(cp, settingsGuid);
                if ((result == null)) {
                    // 
                    // -- create default content
                    result = DesignBlockBaseModel.addDefault<SampleModel>(cp);
                    result.name = tableMetadata.contentName + " " + result.id;
                    result.ccguid = settingsGuid;
                    result.fontStyleId = 0;
                    result.themeStyleId = 0;
                    result.padTop = false;
                    result.padBottom = false;
                    result.padRight = false;
                    result.padLeft = false;
                    // 
                    // -- create custom content
                    result.imageFilename = string.Empty;
                    result.imageAspectRatioId = 3;
                    result.headline = "Lorem Ipsum Dolor";
                    result.description = "<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>";
                    result.embed = string.Empty;
                    result.buttonUrl = string.Empty;
                    result.buttonText = string.Empty;
                    // 
                    result.save(cp);
                    // 
                    // -- track the last modified date
                    cp.Content.LatestContentModifiedDate.Track(result.modifiedDate);
                }
                return result;
            }
        }
    }
}
