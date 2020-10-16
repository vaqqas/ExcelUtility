using Vqs.Excel;

namespace Vqs.ExcelTest.Models.Horizontal
{
    public class TeamMap : ExcelMap<Team>
    {
        public override int Header => 0;

        public static TeamMap Create()
        {
            var map = new TeamMap();

            var type = typeof(Team);
            map.Mapping.Add(1, type.GetProperty(nameof(Team.Name)));
            map.Mapping.Add(2, type.GetProperty(nameof(Team.Designation)));
            map.Mapping.Add(3, type.GetProperty(nameof(Team.DOB)));
            map.Mapping.Add(3, type.GetProperty(nameof(Team.Points)));

            return map;
        }
    }
}