
using System;
using Contensive.Addons.aoFileManager2.Models.View;
using Contensive.Addons.aoFileManager2.Models.Db;
using Contensive.BaseClasses;
using Contensive.Addons.aoFileManager2.Controllers;

namespace Models.View {
    public class SampleViewModel : DesignBlockViewBaseModel {
        // 
        public string imageUrl { get; set; }
        public string headlineTopPadClass { get; set; }
        public string headline { get; set; }
        public string descriptionTopPadClass { get; set; }
        public string embed { get; set; }
        public string embedTopPadClass { get; set; }
        public string description { get; set; }
        public string buttonTopPadClass { get; set; }
        public string buttonText { get; set; }
        public string buttonUrl { get; set; }
        public bool manageAspectRatio { get; set; }
        public string styleAspectRatio { get; set; }
        public string btnStyleSelector { get; set; }
        // 
        // ====================================================================================================
        /// <summary>
        /// Populate the view model from the entity model
        /// </summary>
        /// <param name="cp"></param>
        /// <param name="settings"></param>
        /// <returns></returns>
        public static SampleViewModel create(CPBaseClass cp, SampleModel settings) {
            try {
                // 
                // -- base fields
                var result = create<SampleViewModel>(cp, settings);
                // 
                // -- custom
                result.imageUrl = string.IsNullOrEmpty(settings.imageFilename) ? "" : (cp.Site.FilePath + settings.imageFilename).Replace(" ", "%20");
                result.styleAspectRatio = DesignBlockController.getAspectRationStyle(settings.imageAspectRatioId);
                result.manageAspectRatio = !string.IsNullOrEmpty(result.styleAspectRatio);
                // 
                bool isTopElement = string.IsNullOrWhiteSpace(result.imageUrl);
                result.headline = settings.headline;
                result.headlineTopPadClass = isTopElement & (!string.IsNullOrEmpty(result.headline)) ? "" : "pt-3";
                // 
                isTopElement = isTopElement & string.IsNullOrWhiteSpace(result.headline);
                result.embed = settings.embed;
                result.headlineTopPadClass = isTopElement ? "" : "pt-3";
                // 
                isTopElement = isTopElement & string.IsNullOrWhiteSpace(result.embed);
                result.description = settings.description;
                result.descriptionTopPadClass = isTopElement ? "" : "pt-3";
                // 
                isTopElement = isTopElement & string.IsNullOrWhiteSpace(result.description);
                result.buttonUrl = GenericController.verifyProtocol(settings.buttonUrl);
                result.buttonText = string.IsNullOrWhiteSpace(settings.buttonText) ? string.Empty : settings.buttonText;
                result.btnStyleSelector = settings.btnStyleSelector;
                if ((string.IsNullOrEmpty(result.buttonText) | string.IsNullOrEmpty(result.buttonUrl))) {
                    result.buttonText = "";
                    result.buttonUrl = "";
                }
                result.buttonTopPadClass = isTopElement ? "" : "pt-3";
                // 
                return result;
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                return null;
            }
        }
    }
}
