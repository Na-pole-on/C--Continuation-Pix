using BR.Core.Attributes;
using Activities.Custom.Properties;
using System.Data;
using System.Text.RegularExpressions;
using Activities.Custom;

namespace Namespace_Custom
{
    [LocalizableScreenName(nameof(Resources.GetAddressesUsingRegex_ScreenName), typeof(Resources))]
    [LocalizableRepresentation(nameof(Resources.Custom_Representation), typeof(Resources))]
    [BR.Core.Attributes.Path("Custom")]
    [Image(typeof(GetAddressesUsingRegex), "Activities.Custom.search.png")]
    public class GetAddressesUsingRegex : BR.Core.Activity
    {
        [LocalizableScreenName(nameof(Resources.in_str_path_ScreenName), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.in_str_path_Description), typeof(Resources))]
        [IsFilePathChooser]
        [IsRequired]
        public System.String in_str_path {get; set;} 
        
        [LocalizableScreenName(nameof(Resources.in_str_pattern_ScreenName), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.in_str_pattern_Description), typeof(Resources))]
        [IsRequired]
        public System.String in_str_pattern {get; set;} 
        
        [LocalizableScreenName(nameof(Resources.in_str_range_ScreenName), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.in_str_range_Description), typeof(Resources))]
        [IsRequired]
        public System.String in_str_range {get; set;} 
        
        [LocalizableScreenName(nameof(Resources.in_str_sheet_ScreenName), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.in_str_sheet_Description), typeof(Resources))]
        [IsRequired]
        public System.String in_str_sheet {get; set;} 
        
        [LocalizableScreenName(nameof(Resources.out_list_addresses_ScreenName), typeof(Resources))]
        [LocalizableDescription(nameof(Resources.out_list_addresses_Description), typeof(Resources))]
        [IsOut]
        public List<System.String> out_list_addresses {get; set;} 
        
        public GetAddressesUsingRegex() 
            => out_list_addresses = new List<System.String>();

        public override void Execute(int? optionID)
        {
            DataTableExcel excel = new DataTableExcel(in_str_path, in_str_sheet);
            DataTable dt = excel.GetDataTableRange(in_str_range);

            out_list_addresses.Clear();

            dt.Rows.Cast<DataRow>()
              .Select((DataRow dr) =>
              {
                  for (int j = 1; j <= dr.Table.Columns.Count; j++)
                      if (Regex.IsMatch(dr["Column" + j].ToString(), in_str_pattern))
                          out_list_addresses.Add($"{(char)('A' + j - 1)}{dt.Rows.IndexOf(dr) + 1}");

                  return "";
              }).All((line) =>
              {
                  return true;
              });
        }
    }
}
