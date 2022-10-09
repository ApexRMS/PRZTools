namespace NCC.PRZTools
{
    public class FeatureUV
    {
        public FeatureUV()
        {
        }

        public string GroupHeading { get; set; }

        public string ClassLabel { get; set; }
        
        public string WhereClause { get; set; }

        public int GroupThreshold { get; set; }

        public int GroupGoal { get; set; }

        public int ClassThreshold { get; set; }

        public int ClassGoal { get; set; }

    }
}
