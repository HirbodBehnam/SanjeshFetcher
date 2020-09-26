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
            Tuple.Create("دینی 1", AppendArrays(Range(51,57), Range(69,71))),
            Tuple.Create("دینی 2", AppendArrays(Range(58,63), Range(71,72))),
            Tuple.Create("دینی 3", AppendArrays(Range(64,68), Range(73,75))),
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
            
            Tuple.Create("قضیه تقسیم", new []{105}),
            Tuple.Create("رسم نمودار", new []{106}),
            Tuple.Create("تابع معکوس", new []{107}),
            Tuple.Create("انتقال", new []{108}),
            Tuple.Create("مثلثات", AppendArrays(Range(109,110), new []{121})),
            Tuple.Create("معادله مثلثاتی", new []{111}),
            Tuple.Create("دنباله", new []{112}),
            Tuple.Create("لگاریتم", new []{113}),
            Tuple.Create("حد", Range(114,115)),
            Tuple.Create("پیوستگی", new []{116}),
            Tuple.Create("مجانب", new []{117}),
            Tuple.Create("مشتق", new []{118,120}),
            Tuple.Create("آهنگ تغییرات", new []{119}),
            Tuple.Create("ریاضی 1", Range(105,121)),
            Tuple.Create("حسابان 1", Range(105,121)),
            Tuple.Create("حسابان 2", Range(105,121)),
            Tuple.Create("حسابان", Range(105,121)),
        };
        public static readonly Tuple<string, int[]>[] MathGeometryPositions =
        {
            Tuple.Create("دایره محاطی محیطی", new []{122}), // 11
            Tuple.Create("رسم و مکان هندسی", Range(123,123)), // 10
            Tuple.Create("تالس", Range(124,126)), //TODO: ask; 10
            Tuple.Create("sin/cos", new []{127, 132}),// 11
            Tuple.Create("دایره", new []{128, 131}), // 11
            Tuple.Create("تبدیل و کاربرد", Range(129,129)), // 11
            Tuple.Create("تجسم فضایی", Range(130,130)), // 10
            Tuple.Create("معادله دایره", Range(133,134)), // 12
            Tuple.Create("بیضی", Range(135,135)), // 12
            Tuple.Create("سهمی", Range(136,136)), // 12
            Tuple.Create("ماتریس", Range(137,139)), // 12
            Tuple.Create("هندسه 1", AppendArrays(Range(123,126), Range(130,130), new []{140})),
            Tuple.Create("هندسه 2", AppendArrays(new []{122}, new []{127, 132}, new []{128, 131}, Range(129,129))),
            Tuple.Create("هندسه 3", AppendArrays(Range(133,139))),
            Tuple.Create("هندسه", Range(122,140)), 
        };
        public static readonly Tuple<string, int[]>[] MathGosastePositions =
        {
            Tuple.Create("مجموعه", Range(101,103)),
            Tuple.Create("گزاره", Range(104,104)),
            Tuple.Create("ترکیبیات", Range(141,142)),
            Tuple.Create("لانه کبوتری", Range(143,143)),
            Tuple.Create("احتمال", Range(144,144)),
            Tuple.Create("احتمال شرطی", Range(145,147)),
            Tuple.Create("آمار", Range(148,148)),
            Tuple.Create("قضیه تقسیم", Range(149,149)),
            Tuple.Create("ب.م.م ک.م.م", Range(150,150)),
            Tuple.Create("هم نهشتی", Range(151,151)),
            // 152 is skipped; idk where it is
            Tuple.Create("گراف", new []{153,155}),
            Tuple.Create("احاطه گری", Range(154,154)),
            Tuple.Create("ریاضی 1", AppendArrays(new []{141,144})),
            Tuple.Create("آمار و احتمال", AppendArrays(Range(101,104), Range(145,148))),
            Tuple.Create("گسسته (دوازدهم)", AppendArrays(Range(142,143), Range(149,155))),
            Tuple.Create("گسسته (کل امتحان)", AppendArrays(Range(101,104), Range(141,155))),
        };
        public static readonly Tuple<string, int[]>[] MathPhysicsPositions =
        {
            Tuple.Create("مغناطیس", AppendArrays(new []{156})),
            Tuple.Create("حرکت شناسی", Range(157,161)),
            Tuple.Create("دینامیک", Range(162,165)),
            Tuple.Create("حرکت دایره ای", Range(166,166)),
            Tuple.Create("نوسان", Range(167,169)),
            Tuple.Create("صوت", Range(170,170)),
            Tuple.Create("موج", Range(171,171)),
            Tuple.Create("بازتاب", Range(172,172)),
            Tuple.Create("شکست نور", Range(173,173)),
            Tuple.Create("تداخل امواج", Range(174,174)),
            Tuple.Create("فیزیک اتمی", Range(175,176)),
            Tuple.Create("الکتریسیته ساکن", Range(177,178)),
            Tuple.Create("میدان الکتریکی", Range(179,179)),
            Tuple.Create("خازن", Range(180,180)),
            Tuple.Create("توان", Range(181,181)),
            Tuple.Create("جاری", AppendArrays(Range(182,184),new []{186})),
            Tuple.Create("مغناطیس", new []{185}),
            Tuple.Create("شار", Range(187,188)),
            Tuple.Create("اندازه گیری", Range(189,189)),
            Tuple.Create("کار و انرژی", Range(190,191)),
            Tuple.Create("فشار", Range(192,194)),
            Tuple.Create("گرما", Range(195,196)),
            Tuple.Create("ترمودینامیک", Range(197,200)),
            Tuple.Create("فیزیک 1", Range(189,200)),
            Tuple.Create("فیزیک 2", AppendArrays(Range(177,188),new []{156})),
            Tuple.Create("فیریک 3", Range(157,176)),
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
            //TODO: ask 68,70
            Tuple.Create("دینی 1", AppendArrays(Range(51,56), new []{69})),
            Tuple.Create("دینی 2", AppendArrays(Range(57,61), Range(71,72), new []{74})),
            Tuple.Create("دینی 3", AppendArrays(Range(62,67), new []{73,75})),
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
            Tuple.Create("اندازه گیری", Range(206,206)),
            Tuple.Create("حرکت شناسی", Range(207,209)),
            Tuple.Create("دینامیک", Range(210,212)),
            Tuple.Create("نوسان", Range(213,214)),
            Tuple.Create("صوت", Range(215,215)),
            Tuple.Create("موج", new []{216,218}),
            Tuple.Create("شکست نور", Range(217,217)),
            Tuple.Create("فیزیک هسته ای", Range(219,220)),
            Tuple.Create("الکتریسیته ساکن", Range(221,222)),
            Tuple.Create("خازن", Range(223,223)),
            Tuple.Create("توان", new []{225}),
            Tuple.Create("الکتریسیته جاری", Range(224,227)),
            Tuple.Create("مغناطیس", Range(228,229)),
            Tuple.Create("کار و انرژی", Range(230,230)),
            Tuple.Create("فشار", Range(231,233)),
            Tuple.Create("گرما", Range(234,235)),
            Tuple.Create("فیزیک 1", AppendArrays(Range(230,235),new []{206})),
            Tuple.Create("فیزیک 2", Range(221,229)),
            Tuple.Create("فیزیک 3", Range(207,220)),
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
