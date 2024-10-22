using EditorUIExtension;
using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Config;
using UnityEditor;
using UnityEngine;
using UnityEngine.UIElements;

[E_Name("ExcelToJson")]
public class ExcelToJson : BaseEditorVE<ExcelToJson>
{
    private static ExcelToJsonConfig _config;

    [MenuItem("Tools/ExcelToJson")]
    public static void ShowWindow()
    {
        var window = GetWindow<ExcelToJson>();
        window.Show();
    }

    protected override void InitConfig()
    {
        _config = AssetDatabase.LoadAssetAtPath<ExcelToJsonConfig>("Assets/Editor/ExcelToJsonConfig.asset");
        inputPath = _config.inputPath;
        outputPath = _config.outputPath;
        _exclude = _config.excludePath;
    }

    [E_Editor(EType.Input), E_Name("Excel配置文件路径"), E_Wrap]
    private string inputPath;

    [E_Editor(EType.Button), E_Name("选择文件夹")]
    private void SelectInputFolder()
    {
        string res = EditorUtility.OpenFolderPanel("选择Excel配置文件夹", "Tables", "");
        TextField ele = GetElement("inputPath").ElementAt(1) as TextField;
        ele.value = res;
        _config.inputPath = res;
    }

    [E_Editor(EType.Input), E_Name("Json配置文件路径"), E_Wrap]
    private string outputPath;

    [E_Editor(EType.Button), E_Name("选择文件夹")]
    private void SelectOutputFolder()
    {
        string res = EditorUtility.OpenFolderPanel("选择输出配置文件夹", "Assets/Configs", "");
        TextField ele = GetElement("outputPath").ElementAt(1) as TextField;
        ele.value = res;
        _config.outputPath = res;
    }

    [E_Editor(EType.Input), E_DataType(DataType.String)] [E_Name("排除列表，该列表内的路径下文件不会生成数据类")]
    private List<string> _exclude = new List<string>();


    [VE_Box(true, false)] private string _excludeGroup = "exclude_group";

    [E_Editor(EType.Button), E_Name("添加排除项")]
    [VE_Box("exclude_group")]
    private void AddExclude()
    {
        ListAdd("_exclude");
    }

    [E_Editor(EType.Button), E_Name("移除排除项")]
    [VE_Box("exclude_group")]
    private void RemoveExclude()
    {
        ListRemove("_exclude");
    }

    [E_Editor(EType.ListValueListener), E_Name("_exclude")]
    private void ExcludeValueChanged()
    {
        _config.excludePath = _exclude;
    }

    [E_Editor(EType.Button), E_Name("转换")]
    private void Exchange()
    {
        string[] filePaths = Directory.GetFiles(inputPath, "*.*", SearchOption.AllDirectories);

        foreach (string filePath in filePaths)
        {
            Debug.Log("文件路径 "+ filePath);
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables

                    //Debug.Log(result);

                    string res = DataSetToJson(result);

                    string relativePath = Path.GetDirectoryName(filePath.Substring(inputPath.Length + 1)) + "/";
                    string storeName = outputPath + "/" + relativePath + Path.GetFileNameWithoutExtension(filePath) +
                                       ".json";

                    Debug.Log("输出路径 "+ storeName);

                    if (!Directory.Exists(outputPath + "/" + relativePath))
                    {
                        Directory.CreateDirectory(outputPath + "/" + relativePath);
                    }

                    File.WriteAllText(storeName, res);
                }
            }
        }


        string DataSetToJson(DataSet dataSet)
        {
            var table = dataSet.Tables[0];

            string json = "[";

            for (int i = 3; i < table.Rows.Count; i++)
            {
                string lineData = "";
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    var cellValue = table.Rows[i][j];
                    string columnName = "\"" + table.Rows[1][j] + "\"";

                    string str = cellValue?.ToString();

                    if (table.Rows[2][j].ToString() != "string")
                    {
                        string pattern = @"(?<![""'\\\w\s])\-?\d+(?=\s*:)|(?<![""'\\\w\s])\-?\d+(?=\s*})|(?<![""'\\\w\s])\b(?!True\b|False\b)(?!\b\d+\b)(?<![+*])[^,:]+(?=\b(?![^""""""""""""""""]*""""""""""""""""))|(?<![""'\\\w\s])\+\d+(?=\b(?![^""""""""""""""""]*""""""""""""""""))|(?<![""'\\\w\s])\-\d+(?=\b(?![^""""""""""""""""]*""""""""""""""""))";
                    
                        // 添加引号
                        string replacement = @"""$0""";
        
                        // 替换字符串
                        str = Regex.Replace(str, pattern, replacement);
                    
                        str = str.Replace("True", "true").Replace("False", "false");
                    }
                    else
                    {
                        str = "\"" + str + "\"";
                    }

                    if (string.IsNullOrEmpty(str))
                    {
                        lineData += columnName + ":null";
                    }
                    else if (table.Rows[2][j].ToString().IndexOf("List") == 0 || table.Rows[2][j].ToString().IndexOf("HashSet") == 0)
                    {
                        // Handle as JSON array if it looks like one
                        lineData += columnName + ":[" + str + "]";
                    }
                    else
                    {
                        lineData += columnName + ":" + str;
                    }

                    if (j != table.Columns.Count - 1)
                    {
                        lineData += ",";
                    }
                }

                json += "{" + lineData + "}";
                if (i != table.Rows.Count - 1)
                {
                    json += ",";
                }
            }

            json += "]";

            return json;
        }
    }

    [E_Editor(EType.Button), E_Name("生成Json数据类")]
    private void GenerateConfigBean()
    {
        string[] filePaths = Directory.GetFiles(inputPath, "*.*", SearchOption.AllDirectories);

        foreach (string filePath in filePaths)
        {
            if (CheckExclude(filePath)) continue;

            Debug.Log("文件路径 "+ filePath);

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables

                    //Debug.Log(result);

                    string[] names = new string[result.Tables[0].Columns.Count];
                    string[] types = new string[result.Tables[0].Columns.Count];

                    for (int i = 0; i < names.Length; i++)
                    {
                        names[i] = result.Tables[0].Rows[1][i].ToString();
                        types[i] = result.Tables[0].Rows[2][i].ToString();
                    }

                    // 生成 C# 类文件
                    string className = Path.GetFileNameWithoutExtension(filePath) + "Data";
                    string classContent = GenerateClassFromJson(names, types, className);

                    // 写入到文件
                    if (!Directory.Exists(Application.dataPath + "/Script/Configs/"))
                    {
                        Directory.CreateDirectory(Application.dataPath + "/Script/Configs/");
                    }

                    string outputPath = Application.dataPath + "/Script/Configs/" + className + ".cs";
                    File.WriteAllText(outputPath, classContent, Encoding.UTF8);

                    Console.WriteLine($"Class file generated: {outputPath}");
                }
            }
        }
    }

    public static string GenerateClassFromJson(string[] names, string[] types, string className)
    {
        StringBuilder classBuilder = new StringBuilder();
        classBuilder.AppendLine("using System.Collections.Generic;");
        classBuilder.AppendLine("using Game;");
        classBuilder.AppendLine("using UnityEngine;");

        HashSet<string> usings = new HashSet<string>();

        // Use reflection to get the namespace for each type
        // foreach (var typeName in types)
        // {
        //     Type type = Type.GetType("Game." + typeName);
        //     if (type != null && !string.IsNullOrEmpty(type.Namespace))
        //     {
        //         usings.Add($"using {type.Namespace};");
        //     }
        //     else
        //     {
        //         // Handle custom or unknown types if necessary
        //         // usings.Add($"using YourCustomNamespace;");
        //     }
        // }

        // Add using statements to the class
        foreach (var usingStatement in usings)
        {
            classBuilder.AppendLine(usingStatement);
        }

        // Continue generating the class
        classBuilder.AppendLine();
        classBuilder.AppendLine("namespace GameConfig\n{");
        classBuilder.AppendLine($"    public class {className} : BaseConfig");
        classBuilder.AppendLine("    {");

        // Generate properties
        for (int i = 0; i < names.Length; i++)
        {
            string propertyName = names[i];
            string propertyType = types[i];

            if (propertyName == "id") continue;

            classBuilder.AppendLine($"        public {propertyType} {propertyName} {{ get; set; }}");
        }

        classBuilder.AppendLine();
        // Generate Clone method
        classBuilder.AppendLine($"        public {className} Clone()");
        classBuilder.AppendLine("        {");
        classBuilder.AppendLine($"            {className} data = new {className}");
        classBuilder.AppendLine("            {");
        for (int i = 0; i < names.Length; i++)
        {
            string propertyName = names[i];

            if (propertyName == "id") continue;

            classBuilder.AppendLine($"                {propertyName} = {propertyName},");
        }

        classBuilder.AppendLine("            };");
        classBuilder.AppendLine("            return data;");
        classBuilder.AppendLine("        }");

        classBuilder.AppendLine("    }");
        classBuilder.AppendLine("}");

        return classBuilder.ToString();
    }


    public bool IsNumber(string input)
    {
        return int.TryParse(input, out _) || float.TryParse(input, out _) || double.TryParse(input, out _);
    }

    public bool IsBoolean(string input)
    {
        return bool.TryParse(input, out _);
    }

    public bool IsNumberOrBoolean(string input)
    {
        return IsNumber(input) || IsBoolean(input);
    }

    public string ChangeToStringArr(string[] items)
    {
        string res = "[";
        for (int i = 0; i < items.Length; i++)
        {
            res += "\"" + items[i] + "\"";
            if (i != items.Length - 1)
            {
                res += ",";
            }
        }

        res += "]";
        return res;
    }

    /// <summary>
    /// 判断是都排除
    /// </summary>
    /// <param name="path"></param>
    /// <returns></returns>
    private bool CheckExclude(string path)
    {
        path = path.Replace("\\", "/");
        foreach (var exclude in _exclude)
        {
            if (path.IndexOf(exclude) != -1)
            {
                return true;
            }
        }

        return false;
    }
}