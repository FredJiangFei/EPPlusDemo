using System;
using System.Collections.Generic;

public class QuestionDto
{
    public long Id { get; set; }
    public string Text { get; set; }
    public long QuestionSectionId { get; set; }
    public string Section { get; set; }
    public string Comments { get; set; }
    public int? AuditRating { get; set; }
    public int? RankRating { get; set; }
    public DateTime UpdatedDateTime { get; set; }
    public DateTime StartDate { get; set; }
    public DateTime CompleteDate { get; set; }
    public long? AssignAnaswerUserId { get; set; }
    public string AssignAnaswerUserName { get; set; }
    public long? AnswerUserId { get; set; }
    public string Observations { get; set; }
    public string Recommendations { get; set; }
    public virtual List<CitationDto> Citations { get; set; }

}