using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KumIn_WPF
{
    class KumonLevel
    {
        private string subject;
        private string level;

        public KumonLevel(string subject, string level)
        {
            Subject = subject;
            Level = level;
        }

        public KumonLevel(KumonLevel level)
        {
            Subject = level.Subject;
            Level = level.Level;
        }

        public string Subject
        {
            get { return subject; }
            set { subject = value; }
        }

        public string Level
        {
            get { return level; }
            set { level = value.ToUpper(); }
        }

        public static KumonLevel operator++(KumonLevel currentLevel)
        {
            KumonLevel nextLevel = new KumonLevel(currentLevel);
            if (currentLevel.Subject == "Math")
            {
                nextLevel.Subject = "Math";
                string[] mathLevels = new string[] { "7A", "6A", "5A", "4A", "3A", "2A", "A"
                    , "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "W"};
                for (int i = 0; i < mathLevels.Count(); i++)
                {
                    if (currentLevel.Level == mathLevels[i])
                    {
                        nextLevel.Level = mathLevels[i + 1];
                        return nextLevel;
                    }
                }

            }
            else if (currentLevel.Subject == "Reading")
            {
                nextLevel.Subject = "Reading";
                string[] readingLevels = new string[] { "7A", "6A", "5A", "4A", "3A", "2A", "AI"
                    , "AII", "BI", "BII", "CI", "CII", "DI", "DII", "EI", "EII", "FI", "FII"
                    , "GI", "GII", "HI", "HII", "I", "J", "K", "L", "M", "N", "O", "W"};
                for (int i = 0; i < readingLevels.Count(); i++)
                {
                    if (currentLevel.Level == readingLevels[i])
                    {
                        nextLevel.Level = readingLevels[i + 1];
                        return nextLevel;
                    }
                }
            }

            return nextLevel;
        }
    }
}
