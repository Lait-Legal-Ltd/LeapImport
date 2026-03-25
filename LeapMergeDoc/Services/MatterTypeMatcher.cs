using LeapMergeDoc.Models;

namespace LeapMergeDoc.Services
{
    public class MatterTypeMatcher
    {
        private readonly Dictionary<string, (int areaId, int? caseTypeId)> _exactMatches;
        private readonly Dictionary<string, (int areaId, int? caseTypeId)> _keywordMatches;

        public MatterTypeMatcher()
        {
            _exactMatches = InitializeExactMatches();
            _keywordMatches = InitializeKeywordMatches();
        }

        private Dictionary<string, (int, int?)> InitializeExactMatches()
        {
            return new Dictionary<string, (int, int?)>(StringComparer.OrdinalIgnoreCase)
            {
                // Criminal -> Criminal justice (6) -> Criminal Justice (80)
                { "Criminal (ENG)", (6, 80) },
                { "Criminal (All)", (6, 80) },
                { "Motoring Offence", (6, 80) },

                // Immigration -> Immigration (11) -> Various case types
                { "Immigration General", (11, 104) },
                { "Asylum", (11, 45) },
                { "Appeal", (11, 46) },
                { "Leave to Remain", (11, 47) },
                { "Citizenship", (11, 50) },
                { "Sponsorship", (11, 106) },
                { "Visa Application", (11, 107) },

                // Family -> Family and children (9) -> Various case types
                { "Family General", (9, 93) },
                { "Divorce & Dissolution", (9, 66) },
                { "Financial Remedy", (9, 68) },
                { "Private Child Arrangement", (9, 68) },

                // Probate & Wills -> Wills and Probate (22)
                { "(Legacy) Probate", (22, 110) },
                { "Wills & Powers of Attorney", (22, 83) },
                { "Wills", (22, 83) },
                { "Letters of Administration - Estate Administration", (22, 110) },
                { "Estate Dispute", (22, 83) },

                // Residential Conveyancing -> Residential Conveyancing (19)
                { "Purchase", (19, 61) },
                { "Purchase - Buy to Let", (19, 61) },
                { "Purchase - Right to Buy", (19, 61) },
                { "Sale", (19, 62) },
                { "Transfer", (19, 64) },
                { "Conveyancing General", (19, 61) },

                // Commercial Conveyancing -> Commercial Conveyancing (18)
                { "Commercial Property Purchase", (18, 51) },
                { "Commercial Lease", (18, 54) },
                { "Commercial Lease Assignment", (18, 57) },
                { "New Lease", (18, 54) },
                { "Lending", (18, 60) },

                // Employment -> Employment (8)
                { "Employment Dispute", (8, 81) },

                // Litigation -> Civil litigation (2)
                { "Litigation General", (2, 73) },
                { "Contract Dispute", (2, 73) },
                { "Commercial Dispute", (2, 73) },
                { "Leasehold Property Dispute", (2, 70) },
                { "Housing Dispute", (2, 70) },
                { "Tenancy Dispute", (2, 70) },
                { "Tenancy", (2, 70) },
                { "Paying Party", (2, 105) },

                // Corporate/Company -> Company commercial (3)
                { "Corporate Services", (3, 108) },
                { "Corporate General", (3, 87) },

                // Intellectual Property -> Intellectual property and IT (13)
                { "Intellectual Property General", (13, 96) },

                // Education -> Legal aid and access to justice (14)
                { "Education", (14, 109) },

                // Miscellaneous -> Miscellaneous (23)
                { "Miscellaneous", (23, 74) }
            };
        }

        private Dictionary<string, (int, int?)> InitializeKeywordMatches()
        {
            return new Dictionary<string, (int, int?)>(StringComparer.OrdinalIgnoreCase)
            {
                { "criminal", (6, 80) },
                { "immigration", (11, 104) },
                { "asylum", (11, 45) },
                { "sponsorship", (11, 106) },
                { "visa", (11, 107) },
                { "family", (9, 93) },
                { "divorce", (9, 66) },
                { "probate", (22, 110) },
                { "will", (22, 83) },
                { "estate", (22, 110) },
                { "purchase", (19, 61) },
                { "sale", (19, 62) },
                { "conveyancing", (19, 61) },
                { "commercial lease", (18, 54) },
                { "commercial property", (18, 51) },
                { "employment", (8, 81) },
                { "litigation", (2, 73) },
                { "dispute", (2, 73) },
                { "contract", (2, 73) },
                { "corporate", (3, 108) },
                { "intellectual property", (13, 96) },
                { "citizenship", (11, 50) },
                { "lending", (18, 60) },
                { "tenancy", (2, 70) },
                { "housing", (2, 70) },
                { "education", (14, 109) }
            };
        }

        public MatterTypeMatch MatchMatterType(string? customerMatterType)
        {
            if (string.IsNullOrWhiteSpace(customerMatterType))
            {
                return CreateUnmatchedResult();
            }

            // Try exact match first
            if (_exactMatches.TryGetValue(customerMatterType.Trim(), out var exactMatch))
            {
                return new MatterTypeMatch
                {
                    AreaOfPracticeId = exactMatch.areaId,
                    CaseTypeId = exactMatch.caseTypeId,
                    AreaOfPracticeName = GetAreaOfPracticeName(exactMatch.areaId),
                    CaseTypeName = GetCaseTypeName(exactMatch.caseTypeId),
                    ConfidenceScore = 1.0,
                    IsExactMatch = true
                };
            }

            // Try keyword matching
            var normalizedInput = customerMatterType.Trim().ToLower();
            foreach (var kvp in _keywordMatches)
            {
                if (normalizedInput.Contains(kvp.Key.ToLower()))
                {
                    return new MatterTypeMatch
                    {
                        AreaOfPracticeId = kvp.Value.areaId,
                        CaseTypeId = kvp.Value.caseTypeId,
                        AreaOfPracticeName = GetAreaOfPracticeName(kvp.Value.areaId),
                        CaseTypeName = GetCaseTypeName(kvp.Value.caseTypeId),
                        ConfidenceScore = 0.7,
                        IsExactMatch = false
                    };
                }
            }

            return CreateUnmatchedResult();
        }

        public List<MatterTypeMatch> BatchMatchMatterTypes(List<string> matterTypes)
        {
            return matterTypes.Select(mt => MatchMatterType(mt)).ToList();
        }

        private MatterTypeMatch CreateUnmatchedResult()
        {
            return new MatterTypeMatch
            {
                AreaOfPracticeId = 23,
                CaseTypeId = 74,
                AreaOfPracticeName = "Miscellaneous",
                CaseTypeName = "Miscellaneous",
                ConfidenceScore = 0.0,
                IsExactMatch = false
            };
        }

        private string GetAreaOfPracticeName(int areaId)
        {
            var areaNames = new Dictionary<int, string>
            {
                { 1, "Advocacy" },
                { 2, "Civil litigation" },
                { 3, "Company commercial" },
                { 4, "Competition" },
                { 5, "Consumer, debt and insolvency" },
                { 6, "Criminal justice" },
                { 7, "Dispute resolution" },
                { 8, "Employment" },
                { 9, "Family and children" },
                { 10, "Human rights" },
                { 11, "Immigration" },
                { 12, "In-house" },
                { 13, "Intellectual property and IT" },
                { 14, "Legal aid and access to justice" },
                { 15, "Personal injury" },
                { 16, "Planning" },
                { 17, "Private client" },
                { 18, "Commercial Conveyancing" },
                { 19, "Residential Conveyancing" },
                { 20, "Social welfare and housing" },
                { 21, "Tax" },
                { 22, "Wills and Probate" },
                { 23, "Miscellaneous" }
            };

            return areaNames.TryGetValue(areaId, out var name) ? name : "Unknown";
        }

        private string GetCaseTypeName(int? caseTypeId)
        {
            if (!caseTypeId.HasValue) return "General";

            var caseTypeNames = new Dictionary<int, string>
            {
                { 45, "Asylum" },
                { 46, "Appeal" },
                { 47, "Leave to Remain" },
                { 48, "Entry Clearance" },
                { 49, "Settlement" },
                { 50, "Nationality" },
                { 51, "Purchase" },
                { 52, "Sale" },
                { 53, "Re-Mortgage" },
                { 54, "New Lease Sale" },
                { 55, "New Lease Purchase" },
                { 56, "Licence to Assign" },
                { 57, "Assignment Lease Purchase" },
                { 58, "Assignment Lease Sale" },
                { 59, "Off Licence" },
                { 60, "Lending Service" },
                { 61, "Purchase" },
                { 62, "Sale" },
                { 63, "Re-Mortgage" },
                { 64, "Transfer" },
                { 65, "Registration" },
                { 66, "Divorce" },
                { 67, "Domestic Violence" },
                { 68, "Civil Proceedings" },
                { 69, "Mediation" },
                { 70, "Landlord and Tenant" },
                { 71, "Boundary Dispute" },
                { 72, "Small Claim" },
                { 73, "General Disputes" },
                { 74, "Miscellaneous" },
                { 75, "Lease Extension" },
                { 76, "Second Charge" },
                { 77, "FLR (M)" },
                { 78, "FLR (FP)" },
                { 79, "Lease Surrender" },
                { 80, "Criminal Justice" },
                { 81, "Employment" },
                { 82, "Personal injury" },
                { 83, "Wills and Probate" },
                { 84, "Landlord and Tenant" },
                { 85, "EEA National" },
                { 87, "General" },
                { 93, "General" },
                { 96, "General" },
                { 104, "General" },
                { 105, "General" },
                { 106, "Sponsorship" },
                { 107, "Visa Application" },
                { 108, "Corporate Services" },
                { 109, "Education" },
                { 110, "Estate Administration" }
            };

            return caseTypeNames.TryGetValue(caseTypeId.Value, out var name) ? name : "General";
        }
    }
}
