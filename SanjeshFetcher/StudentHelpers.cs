using System;
using static SanjeshFetcher.Helpers;

namespace SanjeshFetcher
{
    class StudentHelpers
    {
        /// <summary>
        /// Defines the headers of the literature answers for math students
        /// </summary>
        public static readonly Tuple<string, int[]>[] MathLiteraturePositions = 
        {
            Tuple.Create("معنی لغت",Range(1,3)),
            Tuple.Create("املا",Range(4,6)),
            Tuple.Create("تاریخ ادبیات",new []{7}),
            Tuple.Create("آرایه",Range(8,11)),
            Tuple.Create("زبان فارسی",Range(12,16)),
            Tuple.Create("قرابت معنایی",Range(17,25)),
            Tuple.Create("ادبیات فارسی",Range(1,25)),
        };
        /// <summary>
        /// Defines the headers of the arabic answers for math students
        /// </summary>
        public static readonly Tuple<string, int[]>[] MathArabicPositions = 
        {
            Tuple.Create("ترجمه", Range(26,34)),
            Tuple.Create("تعریب", Range(35,35)),
            Tuple.Create("درک مطلب", Range(36,39)),
            Tuple.Create("تحلیل صرفی", Range(40,42)),
            Tuple.Create("ضبط حرکات", Range(43,43)),
            Tuple.Create("قواعد", Range(44,50)),
            Tuple.Create("زبان عربی", Range(26,50)),
            //TODO: more stuff?
        };
        /// <summary>
        /// Defines the headers of the religion answers for math students
        /// </summary>
        public static readonly Tuple<string, int[]>[] MathDiniPositions = 
        {
            //TODO: more stuff
            Tuple.Create("دینی", Range(51,75)),
        };
        /// <summary>
        /// Defines the headers of the english answers for math students
        /// </summary>
        public static readonly Tuple<string, int[]>[] MathEnglishPositions = {
            Tuple.Create("گرامر", Range(76,79)),
            Tuple.Create("لغات", Range(80,87)),
            Tuple.Create("کلوز", Range(88,92)),
            Tuple.Create("درک مطلب اول", Range(93,96)),
            Tuple.Create("درک مطلب دوم", Range(97,100)),
            Tuple.Create("زبان انگلیسی", Range(76,100)),
        };

        /// <summary>
        /// Defines the headers of the calculus answers for math students
        /// </summary>
        public static readonly Tuple<string, int[]>[] MathCalculusPositions =
        {
            //TODO
            Tuple.Create("حسابان", Range(105,121)),
        };
        public static readonly Tuple<string, int[]>[] MathGeometryPositions =
        {
            //TODO
            Tuple.Create("هندسه", Range(122,140)),
        };
        public static readonly Tuple<string, int[]>[] MathGosastePositions =
        {
            //TODO
            Tuple.Create("گسسته", AppendArrays(Range(101,104), Range(141,155))),
        };
        public static readonly Tuple<string, int[]>[] MathPhysicsPositions =
        {
            //TODO
            Tuple.Create("فیزیک", Range(156,200)),
        };
        public static readonly Tuple<string, int[]>[] MathChemistryPositions =
        {
            //TODO
            Tuple.Create("شیمی", Range(200,235)),
        };
        /// <summary>
        /// The main root for the subjects of the math students
        /// first item is for the subject name. The second item is for subjects details
        /// </summary>
        public static readonly Tuple<string, Tuple<string, int[]>[]>[] MathSubjects =
        {
            // omoomi
            Tuple.Create("ادبیات", MathLiteraturePositions),
            Tuple.Create("عربی", MathArabicPositions),
            Tuple.Create("دینی", MathDiniPositions),
            Tuple.Create("زبان", MathEnglishPositions),
            // ekhtesasi
            Tuple.Create("حسابان", MathCalculusPositions),
            Tuple.Create("هندسه", MathGeometryPositions),
            Tuple.Create("گسسته", MathGosastePositions),
            Tuple.Create("فیزیک", MathPhysicsPositions),
            Tuple.Create("شیمی", MathChemistryPositions),
        };

        /// <summary>
        /// Defines the headers of the literature answers for tajrobi students; It is identical to math question positions
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiLiteraturePositions =
        {
            Tuple.Create("معنی لغت",Range(1,3)),
            Tuple.Create("املا",Range(4,6)),
            Tuple.Create("تاریخ ادبیات",new []{7}),
            Tuple.Create("آرایه",Range(8,11)),
            Tuple.Create("زبان فارسی",Range(12,16)),
            Tuple.Create("قرابت معنایی",Range(17,25)),
            Tuple.Create("ادبیات فارسی",Range(1,25)),
        };
        /// <summary>
        /// Defines the headers of the arabic answers for tajrobi students; It's like math students as well
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiArabicPositions =
        {
            Tuple.Create("ترجمه", Range(26,34)),
            Tuple.Create("تعریب", Range(35,35)),
            Tuple.Create("درک مطلب", Range(36,39)),
            Tuple.Create("تحلیل صرفی", Range(40,42)),
            Tuple.Create("ضبط حرکات", Range(43,43)),
            Tuple.Create("قواعد", Range(44,50)),
            Tuple.Create("زبان عربی", Range(26,50)),
            //TODO: more stuff?
        };
        /// <summary>
        /// Defines the headers of the religion answers for tajrobi students
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiDiniPositions =
        {
            //TODO: more stuff
            Tuple.Create("دینی", Range(51,75)),
        };
        /// <summary>
        /// Defines the headers of the english answers for tajrobi students; Just like math students
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiEnglishPositions = {
            Tuple.Create("گرامر", Range(76,79)),
            Tuple.Create("لغات", Range(80,87)),
            Tuple.Create("کلوز", Range(88,92)),
            Tuple.Create("درک مطلب اول", Range(93,96)),
            Tuple.Create("درک مطلب دوم", Range(97,100)),
            Tuple.Create("زبان انگلیسی", Range(76,100)),
        };
        /// <summary>
        /// Defines the headers of the geology questions
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiGeologyPositions =
        {
            //TODO: more stuff
            Tuple.Create("زمین شناسی", Range(101,125)),
        };
        /// <summary>
        /// Defines the headers of the geology questions for tajrobi students
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiMathPositions =
        {
            //TODO: more stuff
            Tuple.Create("ریاضی", Range(126,155)),
        };
        /// <summary>
        /// Defines the headers of the biology questions
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiBiologyPositions =
        {
            //TODO: more stuff
            Tuple.Create("زسیت شناسی", Range(156,205)),
        };
        /// <summary>
        /// Defines the headers of the physics questions for tajrobi students
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiPhysicsPositions =
        {
            //TODO: more stuff
            Tuple.Create("فیزیک", Range(206,235)),
        };
        /// <summary>
        /// Defines the headers of the physics questions for tajrobi students
        /// </summary>
        public static readonly Tuple<string, int[]>[] TajrobiChemistryPositions =
        {
            //TODO: more stuff
            Tuple.Create("شیمی", Range(236,270)),
        };
        /// <summary>
        /// All of the subjects for tajrobi students
        /// </summary>
        public static readonly Tuple<string, Tuple<string, int[]>[]>[] TajrobiSubjects =
        {
            // omoomi
            Tuple.Create("ادبیات", TajrobiLiteraturePositions),
            Tuple.Create("عربی", TajrobiArabicPositions),
            Tuple.Create("دینی", TajrobiDiniPositions),
            Tuple.Create("زبان", TajrobiEnglishPositions),
            // ekhtesasi
            Tuple.Create("زمین شناسی", TajrobiGeologyPositions),
            Tuple.Create("ریاضی", TajrobiMathPositions),
            Tuple.Create("زیست شناسی", TajrobiBiologyPositions),
            Tuple.Create("فیزیک", TajrobiPhysicsPositions),
            Tuple.Create("شیمی", TajrobiChemistryPositions),
        };
        /// <summary>
        /// Total number of questions in math exam
        /// </summary>
        public const int MathQuestions = 235;
        /// <summary>
        /// Total number of tajrobi questions in exam
        /// </summary>
        public const int TajrobiQuestions = 270;
    }
}
