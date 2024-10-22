using System.Collections.Generic;
using Game;
using UnityEngine;

namespace GameConfig
{
    public class TestConfigData : BaseConfig
    {
        public string name { get; set; }
        public BuildItemType type { get; set; }
        public Dictionary<TriggerType,List<EffectData>> variable { get; set; }

        public TestConfigData Clone()
        {
            TestConfigData data = new TestConfigData
            {
                name = name,
                type = type,
                variable = variable,
            };
            return data;
        }
    }
}
