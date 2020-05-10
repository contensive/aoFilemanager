
using Contensive.Models.Db;

namespace Contensive.Addons.aoFileManager2 {
    namespace Models.Db {
        // 
        /// <summary>
        ///     ''' This model provides the common fields for all Design Blocks.
        ///     ''' </summary>
        public class DesignBlockBaseModel : DbBaseModel {
            // 
            // ====================================================================================================
            // -- instance properties
            public string backgroundImageFilename { get; set; }
            public int fontStyleId { get; set; }
            public int themeStyleId { get; set; }
            public bool padTop { get; set; }
            public bool padBottom { get; set; }
            public bool padRight { get; set; }
            public bool padLeft { get; set; }
            public string styleheight { get; set; }
            public bool asFullBleed { get; set; }
        }
    }
}