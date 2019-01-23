using System.Collections.Generic;

public class AuditingQuestionnaireResultDto
{
    public long Id;
    public string ScoreSheet;
    public decimal Complete;
    public decimal Compliance;
    public int Score;
    public int MaxScore;
    public List<AuditingRatingCountDto> AuditingRatingCounts;
}

public class AuditingRatingCountDto
{
    public string Key;
    public int Value;
}