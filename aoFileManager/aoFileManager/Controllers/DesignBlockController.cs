
using System;
using Contensive.BaseClasses;

namespace Contensive.Addons.aoFileManager2 {
    namespace Controllers {
        public static class DesignBlockController {
            // 
            // ====================================================================================================
            /// <summary>
            ///         ''' return the style for each aspectRatioId. if empty, it is not managed
            ///         ''' </summary>
            ///         ''' <param name="imageAspectRatioId"></param>
            ///         ''' <returns></returns>
            public static string getAspectRationStyle(int imageAspectRatioId) {
                switch (imageAspectRatioId) {
                    case 2: {
                            // 
                            // -- 1:1
                            return "designBlockImageAspect-1-1";
                        }

                    case 3: {
                            // 
                            // -- 3:2
                            return "designBlockImageAspect-3-2";
                        }

                    case 4: {
                            // 
                            // -- 4:3
                            return "designBlockImageAspect-4-3";
                        }

                    case 5: {
                            // 
                            // -- 16:9
                            return "designBlockImageAspect-16-9";
                        }

                    case 6: {
                            // 
                            // -- 2:1
                            return "designBlockImageAspect-2-1";
                        }

                    case 7: {
                            // 
                            // -- 3:1
                            return "designBlockImageAspect-3-1";
                        }

                    case 8: {
                            // 
                            // -- 4:1
                            return "designBlockImageAspect-4-1";
                        }

                    case 9: {
                            // 
                            // -- 5:1
                            return "designBlockImageAspect-5-1";
                        }

                    default: {
                            // 
                            // -- as-is or unknown
                            return string.Empty;
                        }
                }
            }
            // 
            // ====================================================================================================
            /// <summary>
            ///         ''' return the instanceId for a design block. It should be an document argument set when the addon is dropped on the page.
            ///         ''' If the addon is created with a json string it should be included as an argument
            ///         ''' If it is not included, the page id is used to make a string
            ///         ''' If no instanceId can be created a blank is returned and should NOT be used.
            ///         ''' If returnHtmlMessage is non-blank, add it to the html
            ///         ''' </summary>
            ///         ''' <param name="cp"></param>
            ///         ''' <param name="designBlockName">A name of the design block. This must be unqiue for each type of design block (i.e. text, tile, etc)</param>
            ///         ''' <param name="returnHtmlMessage"></param>
            ///         ''' <returns>If blank is returned, </returns>
            public static string getSettingsGuid(CPBaseClass cp, string designBlockName, ref string returnHtmlMessage) {
                // 
                // -- check arguments
                if ((string.IsNullOrWhiteSpace(designBlockName)))
                    throw new ApplicationException("getInstanceId called without valid designBlockName.");
                // 
                // -- if this code is running during page rendering, read the instanceId from the process
                string result = cp.Doc.GetText("instanceId");
                if ((!string.IsNullOrWhiteSpace(result)))
                    return result;
                // 
                // -- if this code is running as a remoteMethod, read the instanceId from the form submission
                result = cp.Doc.GetText("forminstanceId");
                if ((!string.IsNullOrWhiteSpace(result)))
                    return result;
                // 
                // -- if there is no instanceId added to the rendering context, try the page Id
                if ((cp.Doc.PageId > 0)) {
                    // 
                    // -- no instance Id, create a unquie string for this page, but display error is already used on this page
                    result = "DesignBlockUsedWithoutInstanceId-[" + designBlockName + "]-PageId-" + cp.Doc.PageId.ToString();
                    if ((!string.IsNullOrEmpty(cp.Doc.GetText(result)))) {
                        // 
                        // -- no instance Id, second occurance, display error
                        returnHtmlMessage += "<p>Error, this design block is used twice on this page. This is only allowed if it was added with the drag-drop tool, or includes a unique instance id.</p>";
                        cp.Site.ErrorReport("Design Block [" + designBlockName + "] on page [#" + cp.Doc.PageId + "," + cp.Doc.PageName + "] does not include an instanceId and was used on the page twice. This is not allowed. To use it twice, used the drag-drop design block tool or manually add the argument \"instanceid\" : \"{unique-guid}\".");
                        return string.Empty;
                    }
                    cp.Doc.SetProperty(result, "used");
                    return result;
                }
                // 
                // -- if this is the admin site, make one up
                if ((cp.Request.PathPage == cp.Site.GetText("adminurl"))) {
                    // 
                    // -- addon run on admin site
                    result = "DesignBlockUsedOnAdminSite-[" + designBlockName + "]";
                    if ((!string.IsNullOrEmpty(cp.Doc.GetText(result)))) {
                        // 
                        // -- admin site, second occurance, display error
                        returnHtmlMessage += "<p>Error, this design block is used twice on the admin site. This is only allowed if it was added with the drag-drop tool, or includes a unique instance id.</p>";
                        cp.Site.ErrorReport("Design Block [" + designBlockName + "] on page [#" + cp.Doc.PageId + "," + cp.Doc.PageName + "] does not include an instanceId and was used on the page twice. This is not allowed. To use it twice, used the drag-drop design block tool or manually add the argument \"instanceid\" : \"{unique-guid}\".");
                        return string.Empty;
                    }
                    return result;
                }
                throw new ApplicationException("Design Block [" + designBlockName + "] called without instanceId must be on a page or the admin site.");
            }
        }
    }
}