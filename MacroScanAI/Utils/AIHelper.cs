/*
 * Copyright (C) 2025 David S. Shelley <davidsmithshelley@gmail.com>
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License 
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace MacroScanAI.Utils
{
    public class ModuleAnalysisResult
    {
        private bool _isExpanded = false;

        [JsonPropertyName("moduleName")]
        public string ModuleName { get; set; } = "";

        [JsonPropertyName("verdict")]
        public string Verdict { get; set; } = "";

        [JsonPropertyName("confidence")]
        public int Confidence { get; set; }

        [JsonPropertyName("summary")]
        public string Summary { get; set; } = "";

        [JsonPropertyName("indicators")]
        public List<string> Indicators { get; set; } = new();
    }

    internal static class AIHelper
    {

        private static bool IsValidJson(string json)
        {
            try
            {
                JsonNode.Parse(json);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string CleanJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return json;

            // All common invisible Unicode characters that break JSON parsing
            var invisible = new char[]
            {
                '\u200B', // zero-width space
                '\u200C', // zero-width non-joiner
                '\u200D', // zero-width joiner
                '\uFEFF'  // zero-width no-break space / BOM
            };

            foreach (var c in invisible)
                json = json.Replace(c.ToString(), "");

            return json;
        }

        private static ModuleAnalysisResult SafeResult(ModuleAnalysisResult? r, string moduleName)
        {
            if (r == null)
            {
                return new ModuleAnalysisResult
                {
                    ModuleName = moduleName,
                    Verdict = "Error",
                    Confidence = 1,
                    Summary = "Failed to parse AI JSON",
                    Indicators = new List<string>(),
                };
            }

            // Make sure none of the properties are null
            r.ModuleName ??= moduleName;
            r.Verdict ??= "Unknown";
            r.Summary ??= "";
            r.Indicators ??= new List<string>();
            return r;
        }

        public static async Task<ModuleAnalysisResult?> AnalyzeVbaModuleAsync(
                                                string apiKey,
                                                string vbaCode,
                                                string moduleName)
        {
            try
            {
                var client = new OpenAI.Chat.ChatClient(
                    model: "gpt-3.5-turbo",
                    apiKey: apiKey
                );

                string prompt = BuildPrompt(moduleName, vbaCode);


                //var completion = await client.CompleteChatAsync(prompt);
                var completion = await client.CompleteChatAsync(prompt).ConfigureAwait(false);
                string json = completion.Value.Content[0].Text;

                // Clean hidden chars before validating
                json = CleanJson(json);

                // If the JSON is malformed → repair it
                if (!IsValidJson(json))
                {
                    json = await FixMalformedJsonAsync(client, json);
                }


                var result = JsonSerializer.Deserialize<ModuleAnalysisResult>(json,
                    new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    });

                result = SafeResult(result, moduleName);


                // Safety fallback (never return null)
                if (result == null)
                {
                    throw new Exception("Failed to parse AI JSON.");
                }

                return result;
            }
            catch (Exception ex)
            {
                return new ModuleAnalysisResult
                {
                    ModuleName = moduleName,
                    Verdict = "Error",
                    Confidence = 1,
                    Summary = $"Exception: {ex.Message}",
                    Indicators = new List<string>()
                };
            }
        }

        static string BuildPrompt(string moduleName, string vbaCode)
        {
            return $@"
You are an expert in detecting malicious VBA in Office documents.

You MUST output exactly one SINGLE JSON object and NOTHING else.

=== REQUIRED JSON SCHEMA (use exactly this structure) ===
{{
  ""moduleName"": ""{moduleName}"",
  ""verdict"": ""Malicious"" | ""Suspicious"" | ""Benign"",
  ""confidence"": <integer 1-100>,
  ""summary"": ""brief explanation"",
  ""indicators"": [
      ""indicator1"",
      ""indicator2""
  ]
}}

RULES:
- Output ONLY the JSON object.
- No explanations.
- No code fences.
- confidence MUST be 1–100 (never 0).
- indicators MUST be a JSON array (possibly empty).
- verdict MUST be exactly: Malicious, Suspicious, or Benign.

Analyze this VBA module:

{vbaCode}
";
        }

        private static async Task<string> FixMalformedJsonAsync(OpenAI.Chat.ChatClient client, string badJson)
        {
            string repairPrompt = $@"
The following text is INVALID JSON. 
Return ONLY corrected valid JSON with no extra text.

Text:
{badJson}
";

            var result = await client.CompleteChatAsync(repairPrompt);

            return result.Value.Content[0].Text;
        }



    }
}
