using Business.Library;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using ClMoneyConverter.Models;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Security.Principal;
using System.Text;
using System;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;

namespace ClMoneyConverter
{
    class Program
    {
        #region 선언
        public static string LocalPcAddr = @"D:\ClMoneyExcel\Data";
        public static bool IsFileFind = false;

        static Excel.Application excelApp = null;
        static Excel.Workbook workBook = null;
        static Excel.Worksheet workSheet = null;

        #endregion


        #region Main
        static void Main()
        {
            #region 중복실행 방지
            if (Common.IsTaskView())
            {
                return;
            }
            #endregion


            Console.WriteLine($"[{string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now)}] Task 시작됨");


            while (true)
            {
                Thread.Sleep(1000);

                DateTime dtNow = DateTime.Now;

                if (!IsFileFind)
                {
                    CheckFolder();
                }
            }
        } 
        #endregion

        #region 파일 찾기
        private static void CheckFolder()
        {
            try
            {
                try
                {
                    var getPath = Path.Combine(LocalPcAddr);

                    var files = Directory.GetFiles(getPath, $"*.xls");

                    var DirectoryPon = files.OrderByDescending(x => x).ToList();
                    if (DirectoryPon != null && DirectoryPon.Count > 0)
                    {
                        var FullPath = DirectoryPon.FirstOrDefault();
                        Console.WriteLine($"[{string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now)}] 파일 찾기 완료 -> [{FullPath}]");
                        IsFileFind = true;

                        DataFileConverter(FullPath);


                        IsFileFind = false;

                    }
                }
                catch (IOException ioexception)
                {
                    Console.WriteLine($"[{string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now)}] [{ioexception}]");
                }
               
            }
            catch (Exception exception)
            {
                Console.WriteLine($"[{string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now)}] [{MethodBase.GetCurrentMethod().Name}] {exception}");
                throw;
            }
        }
        #endregion

        #region 파일 변환
        private static void DataFileConverter(string FullPath)
        {
            try
            {
                string a = "";

                if (FullPath != "")
                {
                    try
                    {
                        #region 선언
                        excelApp = new Excel.Application();                                 // 엑셀 어플리케이션 생성
                        workBook = excelApp.Workbooks.Open(FullPath);                       // 워크북 열기
                        workSheet = workBook.Worksheets.get_Item(1) as Excel.Worksheet;     // 엑셀 첫번째 워크시트 가져오기

                        string str;
                        int rCnt = 0;
                        int cCnt = 0;
                        string sCellData = "";
                        double dCellData;

                        Excel.Range range = workSheet.UsedRange;    // 사용중인 셀 범위를 가져오기 
                        System.Data.DataTable dt = new System.Data.DataTable();

                        #endregion

                        #region 엑셀 변환
                        // 첫 행을 제목으로 
                        for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                        {
                            str = (string)(range.Cells[1, cCnt] as Excel.Range).Value2;
                            dt.Columns.Add(str, typeof(string));
                        }

                        for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                        {
                            string sData = "";
                            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                            {
                                try
                                {
                                    sCellData = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                                    if (sCellData != null)
                                    {
                                        sData += sCellData.ToString() + "|";
                                    }
                                    else
                                    {
                                        sData += "" + "|";
                                    }
                                }
                                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                                {
                                    var _cellData = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                                    sData += _cellData.ToString() + "|";
                                }
                            }
                            sData = sData.Remove(sData.Length - 1, 1);
                            dt.Rows.Add(sData.Split('|'));

                        }
                        

                        workBook.Close(true);   // 워크북 닫기
                        excelApp.Quit();        // 엑셀 어플리케이션 종료
                        #endregion


                        var _Data = Common.ConvertDataTable<T_ClMoney>(dt);

                        var _List_ImportData = new List<T_Import>
                        {
                            new T_Import
                            {
                                C_Date = "날짜",
                                Account = "자산",
                                MainCategory = "대분류",
                                Subcategory = "소분류",
                                History = "내용",
                                Amount = "금액",
                                Category = "수입/지출",
                                Note = "메모",
                            }
                        };


                        #region 데이터 변환 Import 파일 형태로...
                        foreach (var item in _Data)
                        {
                            var _ImportData2 = new T_Import();
                            var _ImportData = new T_Import
                            {
                                C_Date = item.C_Date.Replace(".", "-"),     // 날짜
                                MainCategory = item.MainCategory,           // 대분류
                                Subcategory = item.Subcategory,             // 소분류
                                History = item.History,                     // 내용
                                Amount = item.Amount.Replace("-", "").Replace(",", ""),                       // 금액
                                Note = item.Note,                           // 메모
                            };

                            string _Account = "";
                            if (item.Account.Contains("→"))
                            {
                                var _Account1 = item.Account.Split("→");
                                var _Result1 = GetContains(_Account1[0].Trim());
                                var _Result2 = GetContains(_Account1[1].Trim());
                                _Account = _Result1 + "→" + _Result2;
                            }
                            else if (item.Account.Contains("←"))
                            {
                                var _Account1 = item.Account.Split("←");
                                var _Result1 = GetContains(_Account1[0].Trim());
                                var _Result2 = GetContains(_Account1[1].Trim());
                                _Account = _Result1 + "←" + _Result2;
                            }
                            else
                            {
                                _Account = GetContains(item.Account.Trim());
                            }

                            _ImportData.Account = _Account;


                            if (item.Division == "수입")
                            {
                                #region 수입
                                _ImportData.Category = "수입"; 
                                #endregion
                            }
                            else if (item.Division == "이체")
                            {
                                #region 이체
                                _ImportData.Category = "이체출금";                                

                                if (item.Subcategory == "이체")
                                {
                                    //이체
                                    var data1 = _Account.Split("→");
                                    _ImportData.Account = data1[0].Trim();
                                    _ImportData.MainCategory = data1[1].Trim();
                                    _ImportData.Subcategory = "";
                                }
                                else if (item.Subcategory == "카드대금" || item.Subcategory == "선결제")
                                {
                                    // 카드대금
                                    var data2 = _Account.Split("←");
                                    _ImportData.Account = data2[1].Trim();
                                    _ImportData.MainCategory = data2[0].Trim();
                                    _ImportData.Subcategory = "";
                                    _ImportData.Note = item.Subcategory;
                                } 
                                #endregion
                            }
                            else if (item.Division == "저축")
                            {
                                #region 저축 - 부채상환
                                _ImportData.Category = "이체출금";                                

                                if (item.MainCategory == "부채상환")
                                {
                                    //부채상환
                                    var data1 = _Account.Split("→");
                                    _ImportData.Account = data1[0].Trim();
                                    _ImportData.MainCategory = data1[1].Trim();
                                    _ImportData.Subcategory = "";
                                    _ImportData.Note = item.MainCategory;
                                }
                                else if (item.MainCategory == "저축")
                                {
                                    //부채상환
                                    var data1 = _Account.Split("→");
                                    _ImportData.Account = data1[0].Trim();
                                    _ImportData.MainCategory = data1[1].Trim();
                                    _ImportData.Subcategory = "";
                                    _ImportData.Note = item.Subcategory;
                                }
                                #endregion
                            }
                            else if (item.Division == "")
                            {
                                #region 지출
                                _ImportData.Category = "지출";

                                if (item.Discount != "0")
                                {
                                    _ImportData2.C_Date = _ImportData.C_Date;
                                    _ImportData2.Account = _ImportData.Account;
                                    _ImportData2.MainCategory = "카드할인";
                                    _ImportData2.Subcategory = item.Subcategory;
                                    _ImportData2.History = item.History;
                                    _ImportData2.Category = _ImportData.Category;
                                    _ImportData2.Note = item.History;

                                    _ImportData2.Amount = "-" + Convert.ToString(item.Discount);
                                }
                                #endregion
                            }


                            if (item.History == "카드할인")
                            {
                                if (Convert.ToInt32(item.Amount) > 0)
                                    _ImportData.Amount = "-" + Convert.ToString(item.Amount);
                            }

                            _List_ImportData.Add(_ImportData);

                            if (_ImportData2 != null)
                            {
                                if (_ImportData2.MainCategory == "카드할인")
                                {
                                    _ImportData2.MainCategory = item.MainCategory;
                                    _ImportData2.History = "[카드할인]" + item.History;                                    
                                    _ImportData2.Amount = "-" + Convert.ToString(item.Discount);
                                    _List_ImportData.Add(_ImportData2);
                                }
                            }
                        }
                        #endregion

                        // 파일 저장
                        var _Filename = GetFileName(FullPath);
                        var WriteFilename = $"{LocalPcAddr}\\{_Filename.Replace(".xls", ".tsv")}";
                        WriteToTsv(_List_ImportData, WriteFilename);

                        Console.WriteLine($"[{string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now)}] 파일 저장 완료 !! -> [{WriteFilename}]");


                    }
                    finally
                    {
                        ReleaseObject(workSheet);
                        ReleaseObject(workBook);
                        ReleaseObject(excelApp);
                    }
                }

                // 다 끝난후....
                string src = FullPath;
                string dest = FullPath + "_old";

                File.Copy(src, dest, true);
                File.Delete(src);

                Console.WriteLine($"[{string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now)}] 처리 완료!!");
            }
            catch (Exception exception)
            {
                Console.WriteLine($"[{string.Format("{0:yyyy-MM-dd HH:mm:ss}", DateTime.Now)}] [{MethodBase.GetCurrentMethod().Name}] {exception}");
                throw;
            }
        }
        #endregion


        #region WriteToTsv
        public static void WriteToTsv(IEnumerable<T_Import> _import, string filePath)
        {
            var tsv = string.Join("\n", _import.Select(p => $"{p.C_Date}\t{p.Account}\t{p.MainCategory}\t{p.Subcategory}\t{p.History}\t{p.Amount}\t{p.Category}\t{p.Note}"));
            File.WriteAllText(filePath, tsv);
        }
        #endregion

        #region GetFileName
        public static string GetFileName(string filePath)
        {
            return Path.GetFileName(filePath);
        }
        #endregion

        #region GetContains
        public static string GetContains(string Account)
        {
            var ChangAccountList = new List<T_ChangAccount>
            {
                new T_ChangAccount(){ OldName = "신한 공과금통장Sj", NewName = "Sj 신한 (생활비)"},
                new T_ChangAccount(){ OldName = "삼성카드Sj", NewName = "Sj 삼성카드"},
                new T_ChangAccount(){ OldName = "신한카드Sj", NewName = "Sj 신한카드"},
                new T_ChangAccount(){ OldName = "카카오뱅크Sj-세이프박스", NewName = "카카오뱅크Sj-세이프박스"},
                new T_ChangAccount(){ OldName = "카카오뱅크Sj", NewName = "Sj 카카오"},
                new T_ChangAccount(){ OldName = "토스뱅킹 Sj", NewName = "Sj 토스체크"},
                new T_ChangAccount(){ OldName = "카카오Pay Sj", NewName = "Sj 카카오Pay"},
                new T_ChangAccount(){ OldName = "카카오뱅크cos", NewName = "Cos 카카오"},
                new T_ChangAccount(){ OldName = "COS신한", NewName = "Cos 신한"},
                new T_ChangAccount(){ OldName = "Cos현금", NewName = "COS현금"},
                new T_ChangAccount(){ OldName = "Cos삼성카드", NewName = "Cos 삼성카드"},
                new T_ChangAccount(){ OldName = "삼성카드Cos", NewName = "Cos 삼성카드"},
                new T_ChangAccount(){ OldName = "신한 Cos", NewName = "Cos 신한"},
                new T_ChangAccount(){ OldName = "국민 직장인Sj", NewName = "Sj 국민 (월급)"},
                new T_ChangAccount(){ OldName = "신한 급여통장Sj", NewName = "Sj 신한 (급여)"},
                new T_ChangAccount(){ OldName = "COS농협 아름드리", NewName = "Cos 농협 (아름드리)"},
                new T_ChangAccount(){ OldName = "COS신한체크", NewName = "Cos 신한체크"},
                new T_ChangAccount(){ OldName = "COS농협", NewName = "Cos 농협"},
                new T_ChangAccount(){ OldName = "CNH농협체크", NewName = "Cos 농협체크"},
                new T_ChangAccount(){ OldName = "cos카카오체크", NewName = "Cos 카카오체크"},
                new T_ChangAccount(){ OldName = "sj카카오체크", NewName = "Sj 카카오체크"},
                new T_ChangAccount(){ OldName = "C농협카드", NewName = "Cos 농협카드"},

                new T_ChangAccount(){ OldName = "신한 아름드리 경조사", NewName = "토스 현석이 가족 경조사"},
                new T_ChangAccount(){ OldName = "가족 신한 (경조사)", NewName = "토스 현석이 가족 경조사"},
                new T_ChangAccount(){ OldName = "카카오 (가족모임회비)", NewName = "토스 (Sj 가족모임회비)"},
                new T_ChangAccount(){ OldName = "우리가족모임통장 카카오", NewName = "토스 (Sj 가족모임회비)"},
                new T_ChangAccount(){ OldName = "C농협중앙회(여행)", NewName = "Cos 농협 (소꿈이)"},
                new T_ChangAccount(){ OldName = "C농협(소꿈이)", NewName = "Cos 농협 (소꿈이)"},

                new T_ChangAccount(){ OldName = "소꿈이자유적금(카카오)+", NewName = "소꿈이자유적금(카카오)+"},
                new T_ChangAccount(){ OldName = "소꿈이자유적금(소꿈이)", NewName = "소꿈이자유적금(카카오)+"},
                new T_ChangAccount(){ OldName = "소꿈이자유적금(카카오)", NewName = "소꿈이자유적금(카카오)+"},

                new T_ChangAccount(){ OldName = "우리-부모님용돈", NewName = "Sj 우리 (부모님용돈)"},
                new T_ChangAccount(){ OldName = "신한 FNA 증권거래저축예금", NewName = "Sj 신한 (FNA)"},
                new T_ChangAccount(){ OldName = "사망보험금 카카오뱅크", NewName = "정기예금 카카오뱅크"},
                new T_ChangAccount(){ OldName = "카카오뱅크-현석 세이프박스", NewName = "Hs 카카오 (세이프박스)"},
                new T_ChangAccount(){ OldName = "Sj 카카오-세이프박스", NewName = "카카오뱅크Sj-세이프박스"},

                new T_ChangAccount(){ OldName = "현석-마이홈 청약종합저출", NewName = "현석-마이홈 청약종합저축"},
                new T_ChangAccount(){ OldName = "현석-신한키즈플러스", NewName = "Hs 신한 (키즈플러스)"},

                new T_ChangAccount(){ OldName = "Sj 카카오 (공과금)(5140)", NewName = "Sj 카카오 (공과금)(5140)"},
                new T_ChangAccount(){ OldName = "Sj 카카오 (공과금)", NewName = "Sj 카카오 (공과금)(5140)"},
                new T_ChangAccount(){ OldName = "카카오 Sj 생활비", NewName = "Sj 카카오 (공과금)(5140)"},
                new T_ChangAccount(){ OldName = "카카오 생활비 Sj", NewName = "Sj 카카오 (공과금)(5140)"},
                new T_ChangAccount(){ OldName = "가족 생활비 카카오", NewName = "Sj 카카오 (공과금)(5140)"},

                new T_ChangAccount(){ OldName = "우리생활비통장", NewName = "가족 카카오 (생활비)"},
                new T_ChangAccount(){ OldName = "주택자금대출용 카카오뱅크", NewName = "Sj 토스 주택자금대출용"},
                new T_ChangAccount(){ OldName = "주택자금대출용 토스뱅크", NewName = "Sj 토스 주택자금대출용"},
                new T_ChangAccount(){ OldName = "생활비 Sj", NewName = "Sj 생활비"},
                new T_ChangAccount(){ OldName = "가족생활비", NewName = "Sj 생활비"},
                new T_ChangAccount(){ OldName = "가족생활비 카카오", NewName = "Sj 생활비"},
                new T_ChangAccount(){ OldName = "Sj 생활비 카카오", NewName = "Sj 생활비"},

                new T_ChangAccount(){ OldName = "현석이방 만들기(4534)", NewName = "현석이방 만들기(4534)"},
                new T_ChangAccount(){ OldName = "현석이방 만들기", NewName = "현석이방 만들기(4534)"},

                new T_ChangAccount(){ OldName = "우리생활비 cos", NewName = "Cos 우리생활비"},

                new T_ChangAccount(){ OldName = "Cos 다이로움체크카드", NewName = "Cos 다이로움체크카드"},
                new T_ChangAccount(){ OldName = "Sj 다이로움체크카드", NewName = "Sj 다이로움체크카드"},
                new T_ChangAccount(){ OldName = "Cos 다이로움", NewName = "Cos 다이로움체크카드"},
                new T_ChangAccount(){ OldName = "Sj 다이로움", NewName = "Sj 다이로움체크카드"},

                new T_ChangAccount(){ OldName = "C청약저축", NewName = "C청약저축"},
                new T_ChangAccount(){ OldName = "청약저축", NewName = "청약저축(6672)"},

                new T_ChangAccount(){ OldName = "신한 연금신탁(1422)", NewName = "신한 연금신탁(1422)"},
                new T_ChangAccount(){ OldName = "신한 연금신탁", NewName = "신한 연금신탁(1422)"},
                new T_ChangAccount(){ OldName = "연금저축", NewName = "신한 연금신탁(1422)"},

                new T_ChangAccount(){ OldName = "신한 개인형 IRP(0668)", NewName = "신한 개인형 IRP(0668)"},
                new T_ChangAccount(){ OldName = "신한 개인형 IRP", NewName = "신한 개인형 IRP(0668)"},

                new T_ChangAccount(){ OldName = "재형저축(5004)", NewName = "정기예금 카카오 카드값"},
                new T_ChangAccount(){ OldName = "재형저축", NewName = "재형저축"},

                new T_ChangAccount(){ OldName = "Sj 토스 보관함 적금", NewName = "Sj 토스-맥북구매용"},
                new T_ChangAccount(){ OldName = "Sj 토스 보관함 카드값", NewName = "Sj 토스 보관함 카드값"},
                new T_ChangAccount(){ OldName = "Sj 토스 보관함", NewName = "Sj 토스 보관함 카드값"},

                new T_ChangAccount(){ OldName = "현석가족 모임통장(토스)", NewName = "토스 현석이가족 모임통장"},
                new T_ChangAccount(){ OldName = "현석가족 모임통장", NewName = "토스 현석이가족 모임통장"},

                new T_ChangAccount(){ OldName = "카카오 정기예금", NewName = "정기예금 카카오 카드값"},
            };

            var _Account = Account;
            foreach (var item in ChangAccountList)
            {
                if (_Account.Contains(item.OldName))
                {
                    return _Account.Replace(item.OldName, item.NewName);
                }
            }

            return _Account;
        }
        #endregion

        #region ReleaseObject
        /// <summary>
        /// 액셀 객체 해제 메소드
        /// </summary>
        /// <param name="obj"></param>
        private static void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);  // 액셀 객체 해제
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();   // 가비지 수집
            }
        } 
        #endregion

    }
}
