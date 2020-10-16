using Vqs.Excel;

namespace Vqs.ExcelTest.Models.Vertical
{
    public class VTeamMap : ExcelMap<VTeam>
    {
        public override ExcelMappingDirection MappingDirection => ExcelMappingDirection.Vertical;

        public override int Header => 0;

        public static VTeamMap Create()
        {
            var map = new VTeamMap();

            var type = typeof(VTeam);
            map.Mapping.Add(1, type.GetProperty(nameof(VTeam.Name)));
            map.Mapping.Add(2, type.GetProperty(nameof(VTeam.Designation)));
            map.Mapping.Add(3, type.GetProperty(nameof(VTeam.DOB)));
            map.Mapping.Add(3, type.GetProperty(nameof(VTeam.Points)));

            return map;
        }
    }
}
