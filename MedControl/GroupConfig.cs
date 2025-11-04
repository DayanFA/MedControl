namespace MedControl
{
    public enum GroupMode
    {
        Solo,
        Host,
        Client
    }

    public static class GroupConfig
    {
        public static GroupMode Mode
        {
            get
            {
                var v = Database.GetConfig("group_mode") ?? "solo";
                return v switch { "host" => GroupMode.Host, "client" => GroupMode.Client, _ => GroupMode.Solo };
            }
            set
            {
                Database.SetConfig("group_mode", value == GroupMode.Host ? "host" : value == GroupMode.Client ? "client" : "solo");
            }
        }

        public static string GroupName
        {
            get => Database.GetConfig("sync_group") ?? "default";
            set => Database.SetConfig("sync_group", value ?? "default");
        }

        public static string HostAddress
        {
            get => Database.GetConfig("group_host") ?? string.Empty; // formato: host:49383
            set => Database.SetConfig("group_host", value ?? string.Empty);
        }

        public static int HostPort
        {
            get
            {
                var s = Database.GetConfig("group_port");
                return int.TryParse(s, out var p) ? p : 49383;
            }
            set => Database.SetConfig("group_port", value.ToString());
        }
    }
}
