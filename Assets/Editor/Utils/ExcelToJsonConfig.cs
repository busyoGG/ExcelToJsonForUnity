using System.Collections;
using System.Collections.Generic;
using UnityEngine;

namespace Config
{
    [CreateAssetMenu(fileName = "ExcelToJsonConfig", menuName = "GenerateSO/Config/ExcelToJsonConfig")]

    public class ExcelToJsonConfig : ScriptableObject
    {
        public string inputPath;

        public string outputPath;

        public List<string> excludePath;
    }

}