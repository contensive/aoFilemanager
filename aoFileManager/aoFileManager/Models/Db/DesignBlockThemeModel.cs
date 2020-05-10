
using Contensive.Models.Db;

namespace Contensive.Addons.aoFileManager2 {
    namespace Models.Db {
        // 
        /// <summary>
        ///     ''' This model provides the common fields for all Design Blocks.
        ///     ''' </summary>
        public class DesignBlockThemeModel : DbBaseModel {
            // 
            // ====================================================================================================
            /// <summary>
            /// metadata
            /// </summary>
            public static readonly DbBaseTableMetadataModel tableMetadata = new DbBaseTableMetadataModel("Design Block Themes", "dbthemes", "default", false);
            // 
            // ====================================================================================================
            // -- instance properties
        }
    }
}