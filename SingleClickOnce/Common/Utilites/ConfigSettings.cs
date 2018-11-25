using System;
using System.Configuration;

namespace Common
{
    public abstract class ConfigSettings
    {
        const string ENVIRONMENT = "environment";

        public static Enums.Environments Environment
        {
            get
            {
                try
                {
                    Enums.Environments environment;
                    string value = ConfigurationManager.AppSettings[ENVIRONMENT];

                    if (!Enum.TryParse<Enums.Environments>(value, out environment))
                    {
                        environment = Enums.Environments.NOTSET;
                    }
                    return environment;
                }
                catch
                {
                    throw new SystemException("AppConfig:: Environment key required!");
                }
            }
        }
    }
}
