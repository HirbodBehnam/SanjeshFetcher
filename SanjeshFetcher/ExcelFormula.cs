namespace SanjeshFetcher
{
    class ExcelFormula
    {
        /// <summary>
        /// Calculates the percentage of the last 3 cells in row
        /// </summary>
        public const string SubjectFormula =
            "=(INDIRECT(ADDRESS(ROW(),COLUMN()-3))*3-INDIRECT(ADDRESS(ROW(),COLUMN()-2)))/(INDIRECT(ADDRESS(ROW(),COLUMN()-3))+INDIRECT(ADDRESS(ROW(),COLUMN()-2))+INDIRECT(ADDRESS(ROW(),COLUMN()-1)))/3*100";

        /// <summary>
        /// Calculates the average of the above cells excluding first 2 rows
        /// </summary>
        public const string AverageBottomFormula = "=AVERAGE(INDIRECT(ADDRESS(3,COLUMN())&\":\"&ADDRESS(ROW()-1,COLUMN())))";
        /// <summary>
        /// Used in single question sheets. Used for calculating percentage of the correct answers
        /// </summary>
        public const string QuestionsCorrectPercentage = "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))/(INDIRECT(ADDRESS(ROW()-1,COLUMN()))+INDIRECT(ADDRESS(ROW()+1,COLUMN()))+INDIRECT(ADDRESS(ROW()+3,COLUMN())))*100";
        /// <summary>
        /// Used in single question sheets. Used for calculating percentage of the wrong answers
        /// </summary>
        public const string QuestionsWrongPercentage = "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))/(INDIRECT(ADDRESS(ROW()-3,COLUMN()))+INDIRECT(ADDRESS(ROW()-1,COLUMN()))+INDIRECT(ADDRESS(ROW()+1,COLUMN())))*100";
        /// <summary>
        /// Used in single question sheets. Used for calculating percentage of the white answers
        /// </summary>
        public const string QuestionsWhitePercentage = "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))/(INDIRECT(ADDRESS(ROW()-5,COLUMN()))+INDIRECT(ADDRESS(ROW()-3,COLUMN()))+INDIRECT(ADDRESS(ROW()-1,COLUMN())))*100";
        /// <summary>
        /// Get total percentage of a question
        /// </summary>
        public const string QuestionsPercentage = "=(INDIRECT(ADDRESS(ROW()-6,COLUMN())) * 3 - INDIRECT(ADDRESS(ROW()-4,COLUMN()))) / (INDIRECT(ADDRESS(ROW()-6,COLUMN())) + INDIRECT(ADDRESS(ROW()-4,COLUMN())) + INDIRECT(ADDRESS(ROW()-2,COLUMN()))) / 3 * 100";
    }
}
