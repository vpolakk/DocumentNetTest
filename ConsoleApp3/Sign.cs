using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp3
{
    class Sign
    {
        /// <summary>
        /// Задает или возвращает время подписания документа.
        /// </summary>
        public string SigningTime { get; set; }

        /// <summary>
        /// Задает или возвращает время заверения подписи документа.
        /// </summary>
        public string SignatureTimeStampTime { get; set; }

        /// <summary>
        /// Задает или возвращает наименование подписанта организации.
        /// </summary>
        public string OrganizationName { get; set; }

        /// <summary>
        /// Задает или возвращает имя ответственного лица.
        /// </summary>
        public string Employee { get; set; }

        /// <summary>
        /// Задает или возвращает имя серийный номер сертификата.
        /// </summary>
        public string SerialNumber { get; set; }

        /// <summary>
        /// Задает или возвращает имя период действия сертификата.
        /// </summary>
        public string ValidityPeriod { get; set; }
    }
}
