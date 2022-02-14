namespace D365FOSecurityComparison
{
    public class SecurityComparison
    {
        public string Name { get; set; }
        public LayerType Type { get; set; }
        public Action Comparison { get; set; }
    }
}
