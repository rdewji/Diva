Option Explicit On
Option Strict On
Option Compare Text
Imports DivaComponent
Imports System.Collections.Generic

Public Class ActFuncB
   Inherits DivaCalcTools.BaseCalcObject

#Region "Dims and Declares"
   Dim tmpStr As String = Replace("|Age Last||Age||None|No Change|Monthly|Monthly||Start of Period|" _
                           & "|00 01 02 03 04 05 06 07 08 09 10 11 12 13 14 15 16 17 18 19 20 21|N2|No||Exponential|No|", "|", vbCr)
   Dim tmpStrClasses As String = Replace("||65|||||||||None|||No||No|BANDS||Rate|4%||Rate|0||" _
                              & "|65|||||||||None|||No||No|BANDS||Rate|4%||Rate|0||" _
                              & "|65|||||||||None|||No||No|BANDS||Rate|4%||Rate|0||", "|", vbCr)
   Dim tmpStr2 As String = vbTab & vbCr & vbTab & vbCr & vbTab
   Dim vParamArray() As String = {tmpStr, tmpStrClasses, tmpStr2, tmpStr2}
   Dim ParameterInterestVector, ParameterInterestExpr, ParameterGteeYrs As String

   Dim vEleNamesBase() As String = {"NumLivesStrt_GC", "NumLivesEnd_GC", "NumDeaths_GC", _
                            "NumLivesStrt_SY", "NumLivesEnd_SY", "NumDeaths_SY", _
                            "NumLivesStrt_OT", "NumLivesEnd_OT", "NumDeaths_OT", _
                            "TotalCF_NRA_GC", "TotalCF_NRA_SY", "TotalCF_NRA_OT", _
                            "TotalCF_MaxCV_GC", "TotalCF_MaxCV_SY", "TotalCF_MaxCV_OT", _
                            "TotalPV_NRA_GC", "TotalPV_NRA_SY", "TotalPV_NRA_OT", _
                            "TotalPV_MaxCV_GC", "TotalPV_MaxCV_SY", "TotalPV_MaxCV_OT"}

   Dim vClassEleNames() As String = {"CF_NRA_GC", "PV_NRA_GC", "CF_NRA_SY", "PV_NRA_SY", "CF_NRA_OT", "PV_NRA_OT", _
                             "CF_MaxCV_GC", "PV_MaxCV_GC", "CF_MaxCV_SY", "PV_MaxCV_SY", "CF_MaxCV_OT", "PV_MaxCV_OT"}

   Dim vClassEleDescs() As String = {"Projected Cash Flow, Going Concern Basis", "Present Value of Cash Flows, Going Concern Basis", _
                             "Projected Cash Flow, Solvency Basis", "Present Value of Cash Flows, Solvency Basis", _
                             "Projected Cash Flow, Other Basis", "Present Value of Cash Flows, Other Basis", _
                             "Projected Cash Flow for Age of Max CV, Going Concern Basis", "Present Value of Cash Flows for Age of Max CV, Going Concern Basis", _
                             "Projected Cash Flow for Age of Max CV, Solvency Basis", "Present Value of Cash Flows for Age of Max CV, Solvency Basis", _
                             "Projected Cash Flow for Age of Max CV, Other Basis", "Present Value of Cash Flows for Age of Max CV, Other Basis"}
   Const oNumClassEles As Integer = 11
   Const oClassEleCF_GC As Integer = 0
   Const oClassElePV_GC As Integer = 1
   Const oClassEleCF_SY As Integer = 2
   Const oClassElePV_SY As Integer = 3
   Const oClassEleCF_OT As Integer = 4
   Const oClassElePV_OT As Integer = 5
   Const oClassEleBestCF_GC As Integer = 6
   Const oClassEleBestPV_GC As Integer = 7
   Const oClassEleBestCF_SY As Integer = 8
   Const oClassEleBestPV_SY As Integer = 9
   Const oClassEleBestCF_OT As Integer = 10
   Const oClassEleBestPV_OT As Integer = 11

   Dim vEleDescriptionsBase() As String = {"Number of Lives at Start, Going Concern", "Number of Lives at End, Going Concern", "Number of Deaths, Going Concern", _
                                 "Number of Lives at Start, Solvency", "Number of Lives at End, Solvency", "Number of Deaths, Solvency", _
                                 "Number of Lives at Start, Other", "Number of Lives at End, Other", "Number of Deaths, Other", _
                                 "Total CF for NRA, Going Concern", "Total CF for NRA, Solvency", "Total CF for NRA, Other", _
                                 "Total CF for Max CV Age, Going Concern", "Total CF for Max CV Age, Solvency", "Total CF for Max CV Age, Other", _
                                 "Total PV for NRA, Going Concern", "Total PV for NRA, Solvency", "Total PV for NRA, Other", _
                                 "Total PV for Max CV Age, Going Concern", "Total PV for Max CV Age, Solvency", "Total  PV for Max CV Age, Other"}
   Dim vEleNames As String() = CType(vEleNamesBase.Clone, String())
   Dim vEleDescriptions() As String = CType(vEleDescriptionsBase.Clone, String())
   Const sMaxElementsBase As Integer = 20 'Last value of base elements (i.e. last constant below)
   Dim sMaxElements As Integer = sMaxElementsBase
   Const sEleLivesStart_GC As Integer = 0
   Const sEleLivesEnd_GC As Integer = 1
   Const sEleDeaths_GC As Integer = 2
   Const sEleLivesStart_SY As Integer = 3
   Const sEleLivesEnd_SY As Integer = 4
   Const sEleDeaths_SY As Integer = 5
   Const sEleLivesStart_OT As Integer = 6
   Const sEleLivesEnd_OT As Integer = 7
   Const sEleDeaths_OT As Integer = 8
   Const sEleTotalCF_NRA_GC As Integer = 9
   Const sEleTotalCF_NRA_SY As Integer = 10
   Const sEleTotalCF_NRA_OT As Integer = 11
   Const sEleTotalCF_MaxCV_GC As Integer = 12
   Const sEleTotalCF_MaxCV_SY As Integer = 13
   Const sEleTotalCF_MaxCV_OT As Integer = 14
   Const sEleTotalPV_NRA_GC As Integer = 15
   Const sEleTotalPV_NRA_SY As Integer = 16
   Const sEleTotalPV_NRA_OT As Integer = 17
   Const sEleTotalPV_MaxCV_GC As Integer = 18
   Const sEleTotalPV_MaxCV_SY As Integer = 19
   Const sEleTotalPV_MaxCV_OT As Integer = 20

   Dim vEleVal(,) As Double
   Dim vPropertyTabIndex As Integer
   Dim vNodeName As String

   Dim AnnPmtInd(), AnnPmtSumPost() As Double
   Dim LivesInd(), LivesInd2(), LivesSum(,) As Double

   Dim npvAnnIndPost(), npvAnnSumPost(), npvAnnIndNormalAge(), npvAnnSumNormalAge() As Double
   Dim npvAnnClass(,,), cfAnnClass(,,), npvBestAgeClass(,,), cfBestAgeClass(,,) As Double
   Dim UseAnnuVec As Boolean ', UseInsuVec  keyINSUREMOVE
   Dim vIntrMethods As String() = {"Rate", "Vector"}
   Dim vInflMethods As String() = {"Rate", "Vector", "Rate: Once/Yr", "Vector: Once/Yr"}
   Const kMethodRate As Integer = 0
   Const kMethodVector As Integer = 1
   Const kMethodSeriatim As Integer = 2
   Const kMethodSeriatimVector As Integer = 3    '   Interest Rates
   Const kMethodRate1PerYr As Integer = 3 '   Inflation rates
   Const kMethodVector1PerYr As Integer = 4
   Const kMethodSeriatim1PerYr As Integer = 5
   Dim VecMax, CalcFreq As Integer, CalcExponent, AnnBenFreq As Double
   Dim vNumFmt As String
   Dim UseMortPreRetire As Boolean
   Dim UseRetireLoop As Boolean
   'Dim RetireLoopAgeFactors As Double(,)
   Dim LastGteeType As Integer
   Const kLastGteeNone As Integer = 0
   Const kLastGteeDate As Integer = 1
   Const kLastGteeYears As Integer = 2
   Const kLastGteeFixed As Integer = 3

   Dim RetireAgeDate As Integer
   Const kRetireAgeParameter As Integer = 0
   Const kRetireDateSeriatim As Integer = 1
   Const kRetireAgeSeriatim As Integer = 2

   Dim YrsGteeAnn As Double
   '	Dim UsePostRetireAnn As Boolean	'UsePostRetireIns, keyINSUREMOVE
   Dim UniMalePct, UniFemPct As Double
   'Dim AgeBandList(), MaxAgeBandList As Integer
   Dim MortYearsToProject As Integer
   Dim vIntraYearIsExponential As Boolean

   Dim InterestRateCollection As Collections.Specialized.StringCollection
   Dim InflationRateCollection As Collections.Specialized.StringCollection
   Dim classNames As New System.Collections.Generic.Dictionary(Of String, Integer)  'Keys are STRINGS (i.e. class names), values are integers (i.e. class NUMBERS)

   Dim DivaCalc As DivaCalcTools.CalcFunctions, gv As DivaCalcTools.GlobalVariables
   Dim gc As New DivaCalcTools.GridConstants

   Dim vPropertyTabText As String() = {"Parameters", "Classes", "Seriatim In", "Seriatim Out"}
   Const sTabParameters As Integer = 0
   Const sTabClasses As Integer = 1
   Const sTabSeriatimIn As Integer = 2
   Const sTabSeriatimOut As Integer = 3
   Const sNumTabs As Integer = 4

   Dim CashFlowTiming As String() = {"Start of Period", "Middle of Period", "End of Period"}
   Dim vCashFlowAnnuity As Integer
   Const vCashFlowStart As Integer = 0
   Const vCashFlowMiddle As Integer = 1
   Const vCashFlowEnd As Integer = 2

   Dim vAgeBasis As Integer
   '   Must be consistent with list in SetLists
   Const AgeBasisLast As Integer = 0
   Const AgeBasisNearest As Integer = 1
   Const AgeBasisInterpolate As Integer = 2
   Const AgeBasisInterpolateMonths As Integer = 3

   Dim vParamNames As String() = { _
      "MORTALITY BASIS", _
      "Age basis to use", _
      "SERIATIM INPUT", "Age or Date of Birth?", "Date format (e.g. YYYYMMDD or blank)", "Years of gteed Annuity benefits", _
      "Benefits change at Retirement?", "Calculation Frequency", _
      "Annuity Benefit Amount Frequency", "BENEFITS", "Timing of Annuity Payments", _
      "WHICH OUTPUTS", "Columns to include in Seriatim Out", "Seriatim Column format", _
      "Include supplementary outputs?", "File Name for supplementary outputs", "Method of Intra-Year Survival", _
      "Run Multiple Retirement Ages?"}

   Enum kRowsParams
      HdrMortBasis
      AgeBasis
      HdrSeriatim
      AgeOrDate
      DateFmt
      YrsGteedPmt
      BenChangeRetire
      CalcFreq
      AnnFreq
      HdrBenefits
      CashAnnuity
      HdrOutputs
      SeriatimOutCol
      SeriatimFmt
      SupplementaryOut
      SupFileName
      MethodIntraYear
      UseRetireLoop

      LastRow
   End Enum

   Enum kColsSeriatimIn
      NameID
      className
      AgeDOB1
      Gender1
      AgeDOB2
      Gender2
      JS
      SurvPct
      AnnuBenefit
      PostRetAnn
      LastGteeDate
      IntrRate
      MortPct1
      MortPct2
      BenIndex
   End Enum

   Const sMaxParams As Integer = kRowsParams.LastRow - 1
   Const sMaxClassRows As Integer = kRowsClasses.LastRow - 1

   Dim vClassItems As String() = { _
      "Class Name", _
      "GOING CONCERN - Mortality Basis", "Normal Retirement Age", "Mortality Table Male", "Mortality Table Female", "Improvement Table Male", "Improvement Table Female", "Withdrawal Table Male", "Withdrawal Table Female", _
      "Mortality Table Male Base Year", "Mortality Table Female Base Year", "Mortality Percentage to Apply", "   Mortality Percentage Male", _
      "   Mortality Percentage Female", "Use Unisex?", "  Male Percentage for Unisex", "Use Pre-Retirement Mortality?", "Retirement Loop Factors", _
      "GOING CONCERN - Financial Assumptions", "Interest Rate or Vector", "  Interest Rate (can be formula)", _
      "  Interest Vector (no formula)", "Inflation Rate or Vector", "  Inflation Rate (can be formula)", "  Inflation Vector (no formula)", _
      "SOLVENCY - Mortality Basis", "Normal Retirement Age", "Mortality Table Male", "Mortality Table Female", "Improvement Table Male", "Improvement Table Female", "Withdrawal Table Male", "Withdrawal Table Female", _
      "Mortality Table Male Base Year", "Mortality Table Female Base Year", "Mortality Percentage to Apply", "   Mortality Percentage Male", _
      "   Mortality Percentage Female", "Use Unisex?", "  Male Percentage for Unisex", "Use Pre-Retirement Mortality?", "Retirement Loop Factors", _
      "SOLVENCY - Financial Assumptions", "Interest Rate or Vector", "  Interest Rate (can be formula)", _
      "  Interest Vector (no formula)", "Inflation Rate or Vector", "  Inflation Rate (can be formula)", "  Inflation Vector (no formula)", _
      "OTHER - Mortality Basis", "Normal Retirement Age", "Mortality Table Male", "Mortality Table Female", "Improvement Table Male", "Improvement Table Female", "Withdrawal Table Male", "Withdrawal Table Female", _
      "Mortality Table Male Base Year", "Mortality Table Female Base Year", "Mortality Percentage to Apply", "   Mortality Percentage Male", _
      "   Mortality Percentage Female", "Use Unisex?", "  Male Percentage for Unisex", "Use Pre-Retirement Mortality?", "Retirement Loop Factors", _
      "OTHER - Financial Assumptions", "Interest Rate or Vector", "  Interest Rate (can be formula)", _
      "  Interest Vector (no formula)", "Inflation Rate or Vector", "  Inflation Rate (can be formula)", "  Inflation Vector (no formula)"}

   Enum kRowsClasses
      className
      HdrMortBasis_GC
      RetAge_GC
      MortTable1_GC
      MortTable2_GC
      ImprovTable1_GC
      ImprovTable2_GC
      WxTable1_GC
      WxTable2_GC
      MortTableMYear_GC
      MortTableFYear_GC
      MortPctType_GC
      MortPctMale_GC
      MortPctFem_GC
      UseUnisex_GC
      UnisexMalePct_GC
      UsePreRetMort_GC
      RetLoopAgeFactors_GC
      HdrBenefits_GC
      IntrRateOrVec_GC
      IntrRate_GC
      IntrVec_GC
      InflRateOrVec_GC
      InflRate_GC
      InflVec_GC
      HdrMortBasis_SY
      RetAge_SY
      MortTable1_SY
      MortTable2_SY
      ImprovTable1_SY
      ImprovTable2_SY
      WxTable1_SY
      WxTable2_SY
      MortTableMYear_SY
      MortTableFYear_SY
      MortPctType_SY
      MortPctMale_SY
      MortPctFem_SY
      UseUnisex_SY
      UnisexMalePct_SY
      UsePreRetMort_SY
      RetLoopAgeFactors_SY
      HdrBenefits_SY
      IntrRateOrVec_SY
      IntrRate_SY
      IntrVec_SY
      InflRateOrVec_SY
      InflRate_SY
      InflVec_SY
      HdrMortBasis_OT
      RetAge_OT
      MortTable1_OT
      MortTable2_OT
      ImprovTable1_OT
      ImprovTable2_OT
      WxTable1_OT
      WxTable2_OT
      MortTableMYear_OT
      MortTableFYear_OT
      MortPctType_OT
      MortPctMale_OT
      MortPctFem_OT
      UseUnisex_OT
      UnisexMalePct_OT
      UsePreRetMort_OT
      RetLoopAgeFactors_OT
      HdrBenefits_OT
      IntrRateOrVec_OT
      IntrRate_OT
      IntrVec_OT
      InflRateOrVec_OT
      InflRate_OT
      InflVec_OT

      LastRow
   End Enum

   Const OneTwelfth As Double = 0.083333333333333329

   '   Used for formulas
   Dim ParmValues(1)(,) As Double
   Dim CellType(1)(,) As Integer
   Dim ParmNode(1)(,) As DivaCalcTools.BaseCalcObject
   Dim ParmElement(1)(,) As Integer
   Dim FuncExpressions(1)(,) As DivaCalcTools.Structures.CalcSequ

   Dim CalcValues(1)() As Double

   Dim ParmValuesMI(,) As Double
   Dim CellTypeMI(,) As Integer
   Dim ParmNodeMI(,) As DivaCalcTools.BaseCalcObject
   Dim ParmElementMI(,) As Integer
   Dim FuncExpressionsMI(,) As DivaCalcTools.Structures.CalcSequ

   Dim vMortPctType As Integer
   Const kMortPctTypeNone As Integer = 0
   Const kMortPctTypeInput As Integer = 1
   Const kMortPctTypeSeriatim As Integer = 2
   Const kMortPctTypeSeriatimM1M2 As Integer = 3

   Const kJointIsSingle As Integer = -1
   Const kJointReduceEither As Integer = 0
   Const kJointReduceLife1 As Integer = 1
   Const kJointReduceLife2 As Integer = 2

   Dim vRowsParmsToIgnore As Integer() = {kRowsParams.HdrMortBasis, kRowsParams.HdrSeriatim, kRowsParams.HdrBenefits, kRowsParams.HdrOutputs, _
                                 kRowsParams.AgeOrDate, kRowsParams.DateFmt, kRowsParams.AgeBasis, _
                                 kRowsParams.YrsGteedPmt, kRowsParams.BenChangeRetire, kRowsParams.CalcFreq, kRowsParams.AnnFreq, _
                                kRowsParams.SeriatimOutCol, kRowsParams.SeriatimFmt, kRowsParams.CashAnnuity, _
                                 kRowsParams.SupplementaryOut, kRowsParams.SupFileName, kRowsParams.MethodIntraYear, _
                                 kRowsParams.UseRetireLoop}

   Dim vRowsClassesToIgnore As Integer() = { _
           kRowsClasses.className, _
         kRowsClasses.HdrBenefits_GC, kRowsClasses.MortTable1_GC, kRowsClasses.MortTable2_GC, _
      kRowsClasses.ImprovTable1_GC, kRowsClasses.ImprovTable2_GC, kRowsClasses.WxTable1_GC, kRowsClasses.WxTable2_GC, _
      kRowsClasses.MortPctType_GC, kRowsClasses.UseUnisex_GC, kRowsClasses.UsePreRetMort_GC, kRowsClasses.RetLoopAgeFactors_GC, _
      kRowsClasses.IntrRateOrVec_GC, kRowsClasses.IntrVec_GC, kRowsClasses.InflRateOrVec_GC, kRowsClasses.InflVec_GC, _
         kRowsClasses.HdrBenefits_SY, kRowsClasses.MortTable1_SY, kRowsClasses.MortTable2_SY, _
      kRowsClasses.ImprovTable1_SY, kRowsClasses.ImprovTable2_SY, kRowsClasses.WxTable1_SY, kRowsClasses.WxTable2_SY, _
      kRowsClasses.MortPctType_SY, kRowsClasses.UseUnisex_SY, kRowsClasses.UsePreRetMort_SY, kRowsClasses.RetLoopAgeFactors_SY, _
      kRowsClasses.IntrRateOrVec_SY, kRowsClasses.IntrVec_SY, kRowsClasses.InflRateOrVec_SY, kRowsClasses.InflVec_SY, _
         kRowsClasses.HdrBenefits_OT, kRowsClasses.MortTable1_OT, kRowsClasses.MortTable2_OT, _
      kRowsClasses.ImprovTable1_OT, kRowsClasses.ImprovTable2_OT, kRowsClasses.WxTable1_OT, kRowsClasses.WxTable2_OT, _
      kRowsClasses.MortPctType_OT, kRowsClasses.UseUnisex_OT, kRowsClasses.UsePreRetMort_OT, kRowsClasses.RetLoopAgeFactors_OT, _
      kRowsClasses.IntrRateOrVec_OT, kRowsClasses.IntrVec_OT, kRowsClasses.InflRateOrVec_OT, kRowsClasses.InflVec_OT}

   Dim classTableRows As Integer() = { _
         kRowsClasses.MortTable1_GC, kRowsClasses.MortTable2_GC, _
      kRowsClasses.ImprovTable1_GC, kRowsClasses.ImprovTable2_GC, kRowsClasses.WxTable1_GC, kRowsClasses.WxTable2_GC, _
         kRowsClasses.MortTable1_SY, kRowsClasses.MortTable2_SY, _
      kRowsClasses.ImprovTable1_SY, kRowsClasses.ImprovTable2_SY, kRowsClasses.WxTable1_SY, kRowsClasses.WxTable2_SY, _
         kRowsClasses.MortTable1_OT, kRowsClasses.MortTable2_OT, _
      kRowsClasses.ImprovTable1_OT, kRowsClasses.ImprovTable2_OT, kRowsClasses.WxTable1_OT, kRowsClasses.WxTable2_OT}

   Dim vRowsToRecalc As Integer()()

   '   Programming note - initial output columns must match input columns
   Dim vColSeriatimIn As String() = { _
      "Name/ID", "Class Name", "Age or DOB 1", "Gender 1", "Age or DOB 2", "Gender 2", "Joint or Single", "Survivor Pct", _
         "Annuity Benefit", "PostRetire Annuity Amt (if used)", _
         "Last Gtee Date or Years", "Interest Rate", "Mortality Pct 1", "Mortality Pct 2", "Benefit Indexation"}
   Dim vColSeriatimOut As String() = { _
      "Name/ID", "Class Name", "Age or DOB 1", "Gender 1", "Age or DOB 2", "Gender 2", "Joint or Single", "Survivor Pct", _
      "Annuity Benefit", "Annuity Post Retire Amt", _
      "Last Gtee Date", "Interest Rate", "Mortality Pct 1", "Mortality Pct 2", "Benefit Indexation", _
      "GC: NRA Annuity PV", "GC: Max CV Age", "GC: Max CV Annuity PV", _
      "SY: NRA Annuity PV", "SY: Max CV Age", "SY: Max CV Annuity PV", _
      "OT: NRA Annuity PV", "OT: Max CV Age", "OT: Max CV Annuity PV"}
   'Const sColRetireAge As Integer = 9	 '   Index where the Retirement Age/Date is.  Changed in DoInitialize.
   Const sColMaxGteeDate As Integer = 10 ' Index where last Gtee Date/Age is stored.  Changed in DoInitialize.
   Const sColInterest As Integer = 11  '   Index where seriatim interest stored.  Changed in DoInitialize
   Const sMaxInCol As Integer = 14
   Const sMaxOutCol As Integer = 23
   '   Indexes of out-columns
   Const kOutAnnGC As Integer = 15
   Const kOutBestAgeGC As Integer = 16
   Const kOutBestAnnGC As Integer = 17
   Const kOutAnnSY As Integer = 18
   Const kOutBestAgeSY As Integer = 19
   Const kOutBestAnnSY As Integer = 20
   Const kOutAnnOT As Integer = 21
   Const kOutBestAgeOT As Integer = 22
   Const kOutBestAnnOT As Integer = 23

   Dim WarnCount As Integer
   Dim Seriatim(,) As Insured
   Dim QxTable() As MortTable, ImprovTable() As ImprovementTable, WxTable() As WithdrawalTable
   Dim ClassList As MortClass(,)
   Dim SeriatimOut() As String
   Dim SeriatimCol() As Integer, SeriatimOutLine(), OutColHeading As String
   Dim SuppOut()() As String
   Dim SumOuts() As Double
   Dim SeriaMax As Integer
   Dim SeriatimOutTotalSpacer As String
   Dim UseAge As Boolean

   '   Used for DOB format
   Dim yStart, yLen, mStart, mLen, dStart, dLen, DateType, DateLen As Integer
   Dim ValDate As Date, currDate As Date, currDay As Integer
   'ValDate is the valudation date; currDate is the date at each duration in DoCalcs; currDay is the day of year at each duration in DoCalcs
   'currDate and currDay are used for mortality improvement
   Dim IsFirstScenario, HasAnySeriatimOut, HasAnySupOut As Boolean
   Dim TimePeriodsToDump, ScenariosToDump As Integer
   Dim vSupFileName As String, vSupIO As System.IO.StreamWriter


   '  *************************************
   '  *************************************
   '  ADDED AS PART OF CLASS LOGIC
   '  *************************************
   '  *************************************
   Dim ParsedClassDefn(kRowsClasses.LastRow - 1)() As String
   Dim MortTableToIndex As New System.Collections.Generic.Dictionary(Of String, MortTable)
   Dim MortTableRows As Integer() = {kRowsClasses.MortTable1_GC, kRowsClasses.MortTable1_SY, kRowsClasses.MortTable1_OT, _
          kRowsClasses.MortTable2_GC, kRowsClasses.MortTable2_SY, kRowsClasses.MortTable2_OT}


   Dim ImprovTableRows As Integer() = {kRowsClasses.ImprovTable1_GC, kRowsClasses.ImprovTable1_SY, kRowsClasses.ImprovTable1_OT, _
                             kRowsClasses.ImprovTable2_GC, kRowsClasses.ImprovTable2_SY, kRowsClasses.ImprovTable2_OT}
   Dim ImprovTableToIndex As New System.Collections.Generic.Dictionary(Of String, ImprovementTable)

   Dim WxTableToIndex As New System.Collections.Generic.Dictionary(Of String, WithdrawalTable)
   Dim WxTableRows As Integer() = {kRowsClasses.WxTable1_GC, kRowsClasses.WxTable1_OT, kRowsClasses.WxTable1_SY, _
                           kRowsClasses.WxTable2_GC, kRowsClasses.WxTable2_OT, kRowsClasses.WxTable2_SY}

   '  *************************************
   '  *************************************

   Structure MortClass
      Dim RetAge As Double
      Dim MortTableM As MortTable
      Dim MortTableF As MortTable
      Dim ImprovTableM As ImprovementTable
      Dim ImprovTableF As ImprovementTable
      Dim ImprovFactorsM As ImprovementFactors
      Dim ImprovFactorsF As ImprovementFactors
      Dim WxTableM As WithdrawalTable
      Dim WxTableF As WithdrawalTable
      Dim MortTableMYear As Integer
      Dim MortTableFYear As Integer
      Dim UseImprovementM As Boolean
      Dim UseImprovementF As Boolean
      Dim UseWithdrawalM As Boolean
      Dim UseWithdrawalF As Boolean
      Dim MortPctType As Integer
      Dim MortPctM As Double
      Dim MortPctF As Double
      Dim UseUnisex As Boolean
      Dim UniMalePct As Double
      Dim UsePreRetMort As Boolean
      Dim RetLoopAgeFactors As Double(,)
      Dim IntrMethod As Integer
      Dim IntrFctrStart As Double()
      Dim IntrFctrAnn As Double()
      Dim IntrRate As Double
      Dim IntrVec As Double()
      Dim InflMethod As Integer
      Dim InflRate As Double
      Dim InflVec As Double()
      Dim InflFctr As Double()
   End Structure

   Structure Insured
      Dim Input As String()
      Dim classNum As Integer
      Dim Age1 As Double
      Dim Age2 As Double
      Dim DOB1 As Date
      Dim DOB2 As Date
      Dim MortTable1 As MortTable
      Dim MortTable2 As MortTable
      Dim ImprovTable1 As ImprovementTable
      Dim ImprovTable2 As ImprovementTable
      Dim ImprovFactors1 As ImprovementFactors
      Dim ImprovFactors2 As ImprovementFactors
      Dim WxTable1 As WithdrawalTable
      Dim WxTable2 As WithdrawalTable
      Dim JointIndex As Integer
      Dim SurvivorPct As Double
      'Dim InsBenefit As Double keyINSUREMOVE
      '		Dim AnnBenefit As Double
      'Dim InsBenefitPostRetire As Double keyINSUREMOVE
      Dim AnnBenefit As Double
      Dim PremFctr As Double
      Dim RetireDate As Date
      Dim RetireAgeSeriatim As Double
      Dim LastGteeDate As Date
      Dim LastGteeYears As Double
      Dim MortPct1 As Double
      Dim MortPct2 As Double
      Dim InterestIndex As Integer
      Dim InflationIndex As Integer
      Dim Lx1 As Double  '   Used track continuance in Chronological mode
      Dim Lx2 As Double '   Used track continuance in Chronological mode
      Dim AgeBandIndex As Integer
      Dim IsMale As Boolean
      Dim IsMale2 As Boolean 'Is spouse male
      Dim IsNonSmo As Boolean
   End Structure

   Public Structure MortTable
      Dim Name As String
      Dim StartAge As Integer
      Dim EndAge As Integer
      Dim SelectPer As Integer
      Dim MaxSelectAge As Integer
      Dim Qx(,) As Double
   End Structure

   Public Structure ImprovementTable ' Table of RATES
      Dim Name As String
      Dim StartAge As Integer
      Dim EndAge As Integer
      Dim StartYear As Integer
      Dim EndYear As Integer
      Dim Rx(,) As Double   'Raw rates
   End Structure

   Public Structure ImprovementFactors 'Table of FACTORS
      Dim StartAge As Integer
      Dim EndAge As Integer
      Dim BaseYear As Integer
      Dim numYears As Integer
      Dim Ix(,) As Double
   End Structure

   Public Structure WithdrawalTable
      Dim Name As String
      Dim StartAge As Integer
      Dim EndAge As Integer
      Dim SelectPer As Integer
      Dim MaxSelectAge As Integer
      Dim Wx() As Double
   End Structure

#End Region

#Region "DoInitialize and DoCalcs"

   Public Overrides Function DoMakeElements() As Boolean
      Return Initialize(False)
   End Function

   Overrides Function DoInitialize() As Boolean
      Return Initialize(True)
   End Function


   Private Function Initialize(ByVal FullInitialize As Boolean) As Boolean
      Dim k, k2, kTab, kRow As Integer
      Dim splParam(sNumTabs - 1)(), splLine() As String
      IsFirstScenario = True
      Try
         vNodeName = ParentNode.NodeName

         For kTab = 0 To sNumTabs - 1
            splParam(kTab) = Split(vParamArray(kTab), vbCr)
         Next

         '  *****************
         '  Added with class definitions
         '  *****************
         Dim tmpClassA As String() = splParam(sTabClasses)
         For kRow = 0 To kRowsClasses.LastRow - 1
            ParsedClassDefn(kRow) = Split(tmpClassA(kRow), vbTab)
         Next kRow
         ' *********************

         ReDim vEleVal(sMaxElements, gv.TimeSeriesMax)

         'Get list of classes, store in collection
         kRow = 0
         Dim kCol As Integer

         classNames.Clear()                        'Clear contents of collection.
         For kCol = 0 To ParsedClassDefn(kRow).GetUpperBound(0)
            classNames.Add(ParsedClassDefn(kRow)(kCol), kCol + 1)  'Key is the CLASS NAME, value is the CLASS NUMBER. Start Class Numbers at 1 (not 0).
         Next
         'End get list of classes

         '''SET ROWS TO RECALC
         'vRowsToRecalc is now a jagged array: it is an array of row numbers to recalculate for the parameters and classes tabs
         'Redeclare vRowsToRecalc with an outer array of 2 inner arrays
         ReDim vRowsToRecalc(1)
         'Now set the number of elements in each inner array. First for parameters, then for classes.
         vRowsToRecalc(sTabParameters) = New Integer(sMaxParams - vRowsParmsToIgnore.GetUpperBound(0) - 1) {}
         'There are six inputs to recalculate in the classes tab: interest rate and inflation rate for each valulation type.
         vRowsToRecalc(sTabClasses) = New Integer(8) {}

         'Loop through to get all rows to recalc for parameters tab
         k2 = 0
         For k = 0 To sMaxParams
            If Array.IndexOf(vRowsParmsToIgnore, k) = -1 Then
               vRowsToRecalc(sTabParameters)(k2) = k
               k2 += 1
            End If
         Next k

         'Manually set all rows to recalc for classes tab (there are only six of them)
         vRowsToRecalc(1) = {kRowsClasses.IntrRate_GC, kRowsClasses.IntrRate_OT, kRowsClasses.IntrRate_SY, _
                     kRowsClasses.InflRate_GC, kRowsClasses.InflRate_OT, kRowsClasses.InflRate_SY}
         ''END ROWS TO RECALC

         'Loop through tabs and put functions/node references/input value in correct structure.
         'First define which rows to ignore (those with text values should be ignored, e.g.)
         Dim RowsToIgnore(), MaxCol, MaxRow As Integer

         For kTab = 0 To 1
            Select Case kTab
               Case sTabParameters
                  MaxCol = 1
                  MaxRow = sMaxParams
                  RowsToIgnore = vRowsParmsToIgnore 'Take from pre-set list in declarations above
               Case sTabClasses
                  MaxCol = ParsedClassDefn(0).GetUpperBound(0) + 1
                  RowsToIgnore = vRowsClassesToIgnore 'Take from pre-set list in declarations above
                  MaxRow = sMaxClassRows
               Case Else
                  RowsToIgnore = Nothing
            End Select

            DivaCalc.ReadParameterString(vParamArray(kTab), MaxRow, ParmValues(kTab), CellType(kTab), ParmNode(kTab), ParmElement(kTab), _
                FuncExpressions(kTab), vEleVal, False, vEleNames, sMaxElements, MaxCol, False, 1, 0, vNodeName, RowsToIgnore, Nothing)

            For kRow = 0 To CellType(kTab).GetUpperBound(0)
               For kCol = 0 To CellType(kTab).GetUpperBound(1)
                  If Array.IndexOf(vRowsToRecalc(kTab), kRow) = -1 Then
                     If FullInitialize AndAlso CellType(kTab)(kRow, kCol) = DivaCalcTools.GridConstants.sCellTypeFormula Then
                        FireMsgBox("Error: In Node " & vNodeName & ", Input in tab " & kTab & ", Row " & kRow & " Column " & kCol & ", is a formula but must be a constant")
                     End If
                  End If
               Next
            Next

         Next kTab

         Dim numClasses As Integer = CellType(sTabClasses).GetUpperBound(1) 'Actual number of classes
         Dim classNum As Integer
         Dim basis As Integer
         ReDim ClassList(numClasses - 1, 2) 'Array of arrays. Each "Class" actually has three components: GC, SY, and OT

         Select Case splParam(sTabParameters)(kRowsParams.AgeBasis)
            Case "Age Last"
               vAgeBasis = AgeBasisLast
            Case "Age Nearest"
               vAgeBasis = AgeBasisNearest
            Case "Interpolate"
               vAgeBasis = AgeBasisInterpolate
            Case "Interpolate Months"
               vAgeBasis = AgeBasisInterpolateMonths
            Case Else
               FireMsgBox("In '" & vNodeName & "', did not recognize age basis '" & splParam(sTabParameters)(kRowsParams.AgeBasis) & "'." _
                  & vbCrLf & "Age Last assumed.")
         End Select

         Select Case splParam(sTabParameters)(kRowsParams.CalcFreq)
            Case "Annual"
               CalcFreq = 1
            Case "Semi-Annual"
               CalcFreq = 2
            Case "Quarterly"
               CalcFreq = 4
            Case "Monthly"
               CalcFreq = 12
            Case "Semi-Monthly"
               CalcFreq = 24
            Case "Weekly"
               CalcFreq = 52
            Case "Daily"
               CalcFreq = 365
         End Select

         Select Case splParam(sTabParameters)(kRowsParams.AnnFreq)
            Case "Annual"
               AnnBenFreq = 1.0
            Case "Semi-Annual"
               AnnBenFreq = 2.0
            Case "Quarterly"
               AnnBenFreq = 4.0
            Case "Monthly"
               AnnBenFreq = 12.0
            Case "Semi-Monthly"
               AnnBenFreq = 24.0
            Case "Weekly"
               AnnBenFreq = 52.0
            Case "Daily"
               AnnBenFreq = 365.0
         End Select

         If CalcFreq < gv.TimePeriodsPerYear Then
            FireMsgBox("ERROR: In " & vNodeName & ", Calculation Frequency is '" & splParam(sTabParameters)(kRowsParams.CalcFreq) & "' but it must be " _
               & "at least as frequent as the Diva Model setting.", MsgBoxStyle.Critical)
            Initialize = False
            gv.MasterStop = True
            Exit Function
         End If

         '   Project to end of mortality tables or end of projection, whichever is longer
         MortYearsToProject = Math.Max(121, CInt(Math.Ceiling(gv.TimeSeriesMax / gv.TimePeriodsPerYear)))
         VecMax = MortYearsToProject * CalcFreq
         CalcExponent = 1.0 / CalcFreq

         ParameterGteeYrs = ""
         Select Case splParam(sTabParameters)(kRowsParams.YrsGteedPmt).Trim
            Case "None"
               LastGteeType = kLastGteeNone
               YrsGteeAnn = 0.0
               vColSeriatimOut(sColMaxGteeDate) = "Guaranteed"
               ParameterGteeYrs = splParam(sTabParameters)(kRowsParams.YrsGteedPmt).Trim
            Case "Seriatim Date"
               LastGteeType = kLastGteeDate
               vColSeriatimOut(sColMaxGteeDate) = "Guaranteed Date"
            Case "Seriatim Years"
               LastGteeType = kLastGteeYears
               vColSeriatimOut(sColMaxGteeDate) = "Guaranteed Years"
            Case Else
               LastGteeType = kLastGteeFixed
               ParameterGteeYrs = splParam(sTabParameters)(kRowsParams.YrsGteedPmt)
               YrsGteeAnn = DivaCalc.EvaluateMathExpression(splParam(sTabParameters)(kRowsParams.YrsGteedPmt))
               vColSeriatimOut(sColMaxGteeDate) = "Guaranteed Years"
         End Select


         'Determine whether or not to run multiple retirement ages
         If splParam(sTabParameters)(kRowsParams.UseRetireLoop) = "Yes" Then
            UseRetireLoop = True
            'RetireLoopAgeFactors = InterpolateRetireAgeFactors(InterpretRetireAgeFactors(splParam(sTabParameters)(kRowsParams.RetireLoopAgeFactors)))
         Else
            UseRetireLoop = False
         End If

         ReDim AnnPmtInd(VecMax), AnnPmtSumPost(VecMax)

         ReDim LivesInd(VecMax), LivesInd2(VecMax), LivesSum(2, VecMax)
         ReDim npvAnnIndPost(VecMax), npvAnnSumPost(VecMax), npvAnnIndNormalAge(VecMax), npvAnnSumNormalAge(VecMax) ', npvInsIndPost(VecMax), npvInsSumPost(VecMax) keyINSUREMOVE
         ReDim npvAnnClass(numClasses - 1, 2, VecMax), cfAnnClass(numClasses - 1, 2, VecMax) 'vector by number of classes, number of basis, and number of time periods
         ReDim npvBestAgeClass(numClasses - 1, 2, VecMax), cfBestAgeClass(numClasses - 1, 2, VecMax)  'vector by number of classes, number of basis, and number of time periods

         vCashFlowAnnuity = Array.IndexOf(CashFlowTiming, splParam(sTabParameters)(kRowsParams.CashAnnuity).Trim)
         If vCashFlowAnnuity < 0 Then
            FireMsgBox("In '" & vNodeName & "', didn't recognize timing of Annuity payments: '" & splParam(sTabParameters)(kRowsParams.CashAnnuity) & "'.")
            vCashFlowAnnuity = 0
         End If

         UseAge = splParam(sTabParameters)(kRowsParams.AgeOrDate).Trim = "Age"

         If Not gv.HasTimeSeries Then
            If RetireAgeDate = kRetireDateSeriatim OrElse LastGteeType = kLastGteeDate OrElse Not UseAge Then
               FireMsgBox("In node '" & vNodeName & "', the ActFuncs component is using dates but you have not specified that the model has TimeSeries." _
                     & vbCrLf & "   Go the the Model Parameters form and specify time series and dates.")
               gv.MasterStop = True
               Return False
            End If
         End If

         Dim tmpStr As String = splParam(sTabParameters)(kRowsParams.DateFmt).Trim.ToUpper
         If tmpStr = "" Then
            DateType = 1
            DateLen = 1
         Else
            yStart = Math.Max(tmpStr.IndexOf("Y"), tmpStr.IndexOf("A"))
            yLen = Math.Max(tmpStr.LastIndexOf("Y"), tmpStr.LastIndexOf("A")) - yStart + 1
            mStart = tmpStr.IndexOf("M")
            mLen = tmpStr.LastIndexOf("M") - mStart + 1
            dStart = Math.Max(tmpStr.IndexOf("D"), tmpStr.IndexOf("J"))
            dLen = Math.Max(tmpStr.LastIndexOf("D"), tmpStr.LastIndexOf("J")) - dStart + 1
            If yStart < 0 Then
               DateType = 2
            ElseIf mStart < 0 Then
               If yLen >= 4 Then DateType = 3 Else DateType = 4
            ElseIf dStart < 0 Then
               If yLen >= 4 Then DateType = 5 Else DateType = 6
            Else
               If yLen >= 4 Then DateType = 7 Else DateType = 8
            End If
            DateLen = Math.Max(1, Math.Max(Math.Max(yStart + yLen, mStart + mLen), dStart + dLen) - 1)
         End If


         'What's with the first part of the next line of code? ***
         'If splParam.GetUpperBound(0) >= kRowsParams.MethodIntraYear AndAlso splParam(sTabParameters)(kRowsParams.MethodIntraYear).Trim <> "Linear" Then
         If splParam(sTabParameters)(kRowsParams.MethodIntraYear).Trim <> "Linear" Then
            vIntraYearIsExponential = True
         Else
            vIntraYearIsExponential = False
         End If

         '   Interpret which output columns to include

         Dim tmpList As String = splParam(sTabParameters)(kRowsParams.SeriatimOutCol).Trim
         Dim tmpItems As String()
         Do While 0 <= tmpList.IndexOf("  ")
            tmpList = Replace(tmpList, "  ", " ")
         Loop
         If tmpList.Trim = "" Then
            HasAnySeriatimOut = False
            Erase SeriatimCol, SeriatimOutLine, SumOuts
         Else
            HasAnySeriatimOut = True
            tmpItems = Split(tmpList.Trim, " ")
            Dim MaxOut As Integer = tmpItems.GetUpperBound(0)
            ReDim SeriatimCol(MaxOut)
            ReDim SeriatimOutLine(MaxOut), SumOuts(MaxOut)
            OutColHeading = ""
            SeriatimOutTotalSpacer = ""
            For k = 0 To MaxOut
               SeriatimCol(k) = CInt(tmpItems(k))
               OutColHeading &= vColSeriatimOut(SeriatimCol(k)) & vbTab
               If SeriatimCol(k) <= sMaxInCol Then
                  SeriatimOutTotalSpacer = "Total"
               End If
            Next k
         End If

         ''   Interpret ages  for outputs

         'tmpList = splParam(sTabParameters)(kRowsParams.AgeBands).ToUpper

         'If tmpList.Contains(gv.CommentSeparator) Then
         '	tmpList = tmpList.Substring(0, tmpList.IndexOf(gv.CommentSeparator)) '   Remove single quote for comments
         'End If

         'Do While 0 <= tmpList.IndexOf("  ")
         '	tmpList = Replace(tmpList, "  ", " ")
         'Loop
         'If tmpList.Trim = "" Then
         '	MaxAgeBandList = 0
         '	AgeBandList = {0}
         'Else
         '	tmpItems = Split(tmpList.Trim, " ")
         '	Dim AgeColl As New Collection
         '	Dim kUB, kAuto, kMin, kMax, kStep As Integer
         '	kUB = tmpItems.GetUpperBound(0)
         '	k = 0
         '	Do
         '		If k <= kUB - 2 AndAlso tmpItems(k + 1) = "TO" Then
         '			kMin = CInt(tmpItems(k))
         '			kMax = CInt(tmpItems(k + 2))
         '			If k <= kUB - 4 AndAlso tmpItems(k + 3) = "STEP" Then
         '				kStep = CInt(tmpItems(k + 4))
         '				k += 5
         '			Else
         '				kStep = 1
         '				k += 3
         '			End If
         '			For kAuto = kMin To kMax Step kStep
         '				AgeColl.Add(kAuto)
         '			Next kAuto
         '		Else
         '			'   Single item
         '			AgeColl.Add(CInt(tmpItems(k)))
         '			k += 1
         '		End If
         '	Loop While k <= kUB
         '	MaxAgeBandList = AgeColl.Count - 1
         '	ReDim AgeBandList(MaxAgeBandList)
         '	For kAuto = 1 To AgeColl.Count
         '		AgeBandList(kAuto - 1) = CInt(AgeColl(kAuto))
         '	Next kAuto
         '	Array.Sort(AgeBandList)
         'End If

         ''   Check for duplicates in age bands - Added 20080911 - duplicates were causing some items to be missed.
         'Dim HasDuplicates As Boolean
         'Do
         '	HasDuplicates = False
         '	For k = 1 To AgeBandList.GetUpperBound(0)
         '		If AgeBandList(k) = AgeBandList(k - 1) Then
         '			FireMsgBox("In '" & vNodeName & "', Age band list has duplicated value for age " & CStr(AgeBandList(k)) & "." & vbCrLf _
         '			   & "Duplicate removed.", MsgBoxStyle.Information, gv.AppName)
         '			Dim k1 As Integer
         '			For k1 = k To AgeBandList.GetUpperBound(0) - 1
         '				AgeBandList(k1) = AgeBandList(k1 + 1)
         '			Next k1
         '			MaxAgeBandList = AgeBandList.GetUpperBound(0) - 1
         '			ReDim Preserve AgeBandList(MaxAgeBandList)
         '			HasDuplicates = True
         '			Exit For
         '		End If
         '	Next k
         'Loop While HasDuplicates


         Dim tmpTest As String
         If splParam.GetUpperBound(0) >= kRowsParams.SupplementaryOut Then
            tmpTest = splParam(sTabParameters)(kRowsParams.SupplementaryOut).Trim()
            vSupFileName = splParam(sTabParameters)(kRowsParams.SupFileName).Trim
         Else
            tmpTest = "None"
            vSupFileName = ""
         End If

         Select Case tmpTest
            Case "None"
               TimePeriodsToDump = -1
               ScenariosToDump = -1
               HasAnySupOut = False
            Case "Initial TimePeriod, All Scenarios"
               TimePeriodsToDump = 0
               ScenariosToDump = gv.MaxScenarioNumThisThread
               HasAnySupOut = True
            Case "All TimePeriods, First Scenario"
               TimePeriodsToDump = gv.TimeSeriesMax
               ScenariosToDump = 1
               HasAnySupOut = True
            Case "All TimePeriods, All Scenarios"
               TimePeriodsToDump = gv.TimeSeriesMax
               ScenariosToDump = gv.MaxScenarioNumThisThread
               HasAnySupOut = True
            Case Else
               TimePeriodsToDump = -1
               ScenariosToDump = -1
               HasAnySupOut = False
         End Select
         Try
            If TimePeriodsToDump >= 0 Or ScenariosToDump > 0 Then
               If gv.IsRunning AndAlso gv.ThisThreadNum = 1 Then
                  Try
                     vSupIO.Close()
                  Catch ex As Exception
                     '   Nothing
                  End Try
                  vSupIO = New System.IO.StreamWriter(vSupFileName)
                  vSupIO.WriteLine("TimePeriods up to " & Format(TimePeriodsToDump, "N0") & ", Scenarios up to " & Format(ScenariosToDump, "N0"))
               End If
            End If
         Catch ex As Exception
            FireMsgBox("In '" & vNodeName & "', couldn't create supplementary file output '" & vSupFileName & "'" & vbCrLf & ex.Message)
            TimePeriodsToDump = -1
            ScenariosToDump = -1
            HasAnySupOut = False
         End Try

         If HasAnySupOut AndAlso gv.NumThreads > 1 And gv.ThisThreadNum = 1 Then
            FireMsgBox("In '" & vNodeName & "', the seriatim supplementary output will only be picking up the first thread." _
               & vbCrLf & "You can set the model to single thread mode if you need to see all scenarios.", MsgBoxStyle.Information, vNodeName)
         End If
         If gv.ThisThreadNum > 1 Then HasAnySupOut = False


         'Make list of element names
         Dim tmpPrefix As String
         sMaxElements = sMaxElementsBase + (oNumClassEles + 1) * numClasses
         ReDim Preserve vEleNames(sMaxElements)
         ReDim Preserve vEleDescriptions(sMaxElements)
         For kClassEle As Integer = 0 To oNumClassEles
            For kClass As Integer = 0 To numClasses - 1
               tmpPrefix = ParsedClassDefn(kRowsClasses.className)(kClass) & " "
               vEleNames(sMaxElementsBase + kClassEle * (numClasses) + kClass + 1) = tmpPrefix & vClassEleNames(kClassEle)
               vEleDescriptions(sMaxElementsBase + kClassEle * (numClasses) + kClass + 1) = tmpPrefix & ": " & vClassEleDescs(kClassEle)
            Next kClass
         Next kClassEle

         '   Make list of Age-related element names

         'sMaxElements = sMaxElementsBase	'+ oNumAgeEles * (MaxAgeBandList + 1)
         'ReDim Preserve vEleNames(sMaxElements)
         'ReDim Preserve vEleDescriptions(sMaxElements)
         'If vEleVal.GetUpperBound(0) <> sMaxElements OrElse vEleVal.GetUpperBound(1) <> gv.TimeSeriesMax Then
         '	ReDim vEleVal(sMaxElements, gv.TimeSeriesMax)
         'End If
         'Dim kAge, kIndex As Integer, tmpPrefix, tmpPrefixDesc As String
         'For kAge = 0 To MaxAgeBandList
         '	kIndex = 1 + sMaxElementsBase + oNumAgeEles * kAge
         '	tmpPrefix = "A" & Format(AgeBandList(kAge), "00")
         '	If kAge < MaxAgeBandList Then
         '		If AgeBandList(kAge + 1) - AgeBandList(kAge) > 1 Then
         '			tmpPrefixDesc = "Age " & Format(AgeBandList(kAge), "000") & " to " & Format(AgeBandList(kAge + 1) - 1, "000")
         '		Else
         '			tmpPrefixDesc = "Age " & Format(AgeBandList(kAge), "000")
         '		End If
         '	Else
         '		tmpPrefixDesc = "Age " & Format(AgeBandList(kAge), "000") & " and up"
         '	End If
         '	vEleNames(kIndex + oEleInsPmt) = tmpPrefix & vEleByAge(oEleInsPmt)
         '	vEleNames(kIndex + oElePvInsPmt) = tmpPrefix & vEleByAge(oElePvInsPmt)
         '	vEleNames(kIndex + oEleAnnPmt) = tmpPrefix & vEleByAge(oEleAnnPmt)
         '	vEleNames(kIndex + oElePvAnnPmt) = tmpPrefix & vEleByAge(oElePvAnnPmt)
         '	vEleNames(kIndex + oEleInsPmtNS) = tmpPrefix & vEleByAge(oEleInsPmtNS)
         '	vEleNames(kIndex + oElePvInsPmtNS) = tmpPrefix & vEleByAge(oElePvInsPmtNS)
         '	vEleNames(kIndex + oEleAnnPmtNS) = tmpPrefix & vEleByAge(oEleAnnPmtNS)
         '	vEleNames(kIndex + oElePvAnnPmtNS) = tmpPrefix & vEleByAge(oElePvAnnPmtNS)
         '	vEleNames(kIndex + oEleInsPmtSmo) = tmpPrefix & vEleByAge(oEleInsPmtSmo)
         '	vEleNames(kIndex + oElePvInsPmtSmo) = tmpPrefix & vEleByAge(oElePvInsPmtSmo)
         '	vEleNames(kIndex + oEleAnnPmtSmo) = tmpPrefix & vEleByAge(oEleAnnPmtSmo)
         '	vEleNames(kIndex + oElePvAnnPmtSmo) = tmpPrefix & vEleByAge(oElePvAnnPmtSmo)
         '	vEleDescriptions(kIndex + oEleInsPmt) = tmpPrefixDesc & " Insurance Pmt in period"
         '	vEleDescriptions(kIndex + oElePvInsPmt) = tmpPrefixDesc & " pv of Future Ins pmts"
         '	vEleDescriptions(kIndex + oEleAnnPmt) = tmpPrefixDesc & " Annuity pmt in period"
         '	vEleDescriptions(kIndex + oElePvAnnPmt) = tmpPrefixDesc & " pv of Future Ann pmts"
         '	vEleDescriptions(kIndex + oEleInsPmtNS) = tmpPrefixDesc & " Insurance Pmt in period NonSmoker"
         '	vEleDescriptions(kIndex + oElePvInsPmtNS) = tmpPrefixDesc & " pv of Future Ins pmts NonSmoker"
         '	vEleDescriptions(kIndex + oEleAnnPmtNS) = tmpPrefixDesc & " Annuity pmt in period NonSmoker"
         '	vEleDescriptions(kIndex + oElePvAnnPmtNS) = tmpPrefixDesc & " pv of Future Ann pmts NonSmoker"
         '	vEleDescriptions(kIndex + oEleInsPmtSmo) = tmpPrefixDesc & " Insurance Pmt in period Smoker"
         '	vEleDescriptions(kIndex + oElePvInsPmtSmo) = tmpPrefixDesc & " pv of Future Ins pmts Smoker"
         '	vEleDescriptions(kIndex + oEleAnnPmtSmo) = tmpPrefixDesc & " Annuity pmt in period Smoker"
         '	vEleDescriptions(kIndex + oElePvAnnPmtSmo) = tmpPrefixDesc & " pv of Future Ann pmts Smoker"
         'Next kAge


         If Not FullInitialize Then Return True
         '   Remove last tab


         If HasAnySeriatimOut Or HasAnySupOut Then
            OutColHeading = OutColHeading.Substring(0, OutColHeading.Length - 1)
         End If
         vNumFmt = splParam(sTabParameters)(kRowsParams.SeriatimFmt)


         '   Interpret list of Shock Years and factors

         If tmpList.Contains(gv.CommentSeparator) Then
            tmpList = tmpList.Substring(0, tmpList.IndexOf(gv.CommentSeparator)) '   Remove single quote for comments
         End If
         Do While 0 <= tmpList.IndexOf("  ")
            tmpList = Replace(tmpList, "  ", " ")
         Loop
         tmpItems = Split(tmpList.Trim, " ")

         Dim ThisYr As Integer, ThisFctr As Double
         Dim tMin, tMax As Integer
         For k = 0 To tmpItems.GetUpperBound(0) - 1 Step 2
            ThisYr = CInt(tmpItems(k))
            ThisFctr = DivaCalc.EvaluateMathExpression(tmpItems(k + 1))
            Select Case CalcFreq
               Case 1
                  tMin = Math.Max(0, CInt(DateDiff(DateInterval.Year, gv.TimeSeries(0), DateSerial(ThisYr, 1, 1))))
                  tMax = Math.Min(VecMax, CInt(DateDiff(DateInterval.Year, gv.TimeSeries(0), DateSerial(ThisYr, 12, 31))))
               Case 2, 4, 12, 24
                  tMin = Math.Max(0, CInt(DateDiff(DateInterval.Month, gv.TimeSeries(0), DateSerial(ThisYr, 1, 1)) / 12.0 * CalcFreq))
                  tMax = Math.Min(VecMax, CInt(DateDiff(DateInterval.Month, gv.TimeSeries(0), DateSerial(ThisYr, 12, 31)) / 12.0 * CalcFreq))
               Case 52
                  tMin = Math.Max(0, CInt(DateDiff(DateInterval.DayOfYear, gv.TimeSeries(0), DateSerial(ThisYr, 1, 1)) / 7.0))
                  tMax = Math.Min(VecMax, CInt(DateDiff(DateInterval.DayOfYear, gv.TimeSeries(0), DateSerial(ThisYr, 12, 31)) / 7.0))
               Case 365
                  tMin = Math.Max(0, CInt(DateDiff(DateInterval.DayOfYear, gv.TimeSeries(0), DateSerial(ThisYr, 1, 1))))
                  tMax = Math.Min(VecMax, CInt(DateDiff(DateInterval.DayOfYear, gv.TimeSeries(0), DateSerial(ThisYr, 12, 31))))
            End Select

         Next k

         'Read in all tables
         If Not ReadQxSheets() Then Return False
         If Not ReadImprovSheets() Then Return False
         If Not ReadWxSheets() Then Return False

         'Note that calculation of improvement FACTORS is done after class parameters are initialized.
         '(This is because we need the Base Table Year to calculate the factors)

         'Set up collection objects to hold mortality tables with their names as the key.
         MortTableToIndex.Clear()
         ImprovTableToIndex.Clear()
         WxTableToIndex.Clear()

         For kTable As Integer = 0 To QxTable.Length - 1
            MortTableToIndex.Add(QxTable(kTable).Name, QxTable(kTable))
         Next

         For kTable As Integer = 0 To WxTable.Length - 1
            WxTableToIndex.Add(WxTable(kTable).Name, WxTable(kTable))
         Next

         For kTable As Integer = 0 To ImprovTable.Length - 1
            ImprovTableToIndex.Add(ImprovTable(kTable).Name, ImprovTable(kTable))
         Next
         'End Read in tables

         ''' INITIALIZE ALL CLASS PARAMETERS
         ''' Note that classes are stored as an array of 3 "class" structures, one for each valuation basis

         'Loop through all classes

         'Dim tmpVec As Double()

         For classNum = 0 To numClasses - 1
            'For each class, loop through three basis types (GC, SY, OT)
            For basis = 0 To 2
               With ClassList(classNum, basis)
                  Select Case basis
                     Case 0   'GC Basis
                        .MortTableM = MortTableToIndex(ParsedClassDefn(kRowsClasses.MortTable1_GC)(classNum).Trim.ToUpper)
                        .MortTableF = MortTableToIndex(ParsedClassDefn(kRowsClasses.MortTable2_GC)(classNum).Trim.ToUpper)
                        .MortTableMYear = CInt(ParsedClassDefn(kRowsClasses.MortTableMYear_GC)(classNum))
                        .MortTableFYear = CInt(ParsedClassDefn(kRowsClasses.MortTableFYear_GC)(classNum))

                        If ParsedClassDefn(kRowsClasses.WxTable1_GC)(classNum).Trim = "" Then
                           .WxTableM = Nothing
                           .UseWithdrawalM = False
                        Else
                           .WxTableM = WxTableToIndex(ParsedClassDefn(kRowsClasses.WxTable1_GC)(classNum).Trim.ToUpper)
                           .UseWithdrawalM = True
                        End If

                        If ParsedClassDefn(kRowsClasses.WxTable2_GC)(classNum).Trim = "" Then
                           .WxTableF = Nothing
                           .UseWithdrawalF = False
                        Else
                           .WxTableF = WxTableToIndex(ParsedClassDefn(kRowsClasses.WxTable1_GC)(classNum).Trim.ToUpper)
                           .UseWithdrawalF = True
                        End If

                        If ParsedClassDefn(kRowsClasses.ImprovTable1_GC)(classNum).Trim = "" Then
                           .ImprovTableM = Nothing
                           .UseImprovementM = False
                           .ImprovFactorsM = Nothing
                        Else
                           .ImprovTableM = ImprovTableToIndex(ParsedClassDefn(kRowsClasses.ImprovTable1_GC)(classNum).Trim.ToUpper)
                           .UseImprovementM = True
                           .ImprovFactorsM = ComputeIxFactors(.ImprovTableM, .MortTableMYear, Math.Max(MortYearsToProject, 120))
                        End If

                        If ParsedClassDefn(kRowsClasses.ImprovTable2_GC)(classNum).Trim = "" Then
                           .ImprovTableF = Nothing
                           .UseImprovementF = False
                           .ImprovFactorsF = Nothing
                        Else
                           .ImprovTableF = ImprovTableToIndex(ParsedClassDefn(kRowsClasses.ImprovTable2_GC)(classNum).Trim.ToUpper)
                           .UseImprovementF = True
                           .ImprovFactorsF = ComputeIxFactors(.ImprovTableF, .MortTableFYear, Math.Max(MortYearsToProject, 120))
                        End If



                        'Retirement age must be rounded to nearest month in order to match up with ages in the retirement loop
                        .RetAge = Math.Round(DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.RetAge_GC)(classNum)) * 12.0) / 12.0

                        Select Case ParsedClassDefn(kRowsClasses.MortPctType_GC)(classNum)
                           Case "None"
                              .MortPctType = kMortPctTypeNone
                           Case "Inputs"
                              .MortPctType = kMortPctTypeInput
                              .MortPctM = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.MortPctMale_GC)(classNum))
                              .MortPctF = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.MortPctFem_GC)(classNum))
                           Case "Seriatim"
                              .MortPctType = kMortPctTypeSeriatim
                           Case "Seriatim [M1 * Qx + M2]"
                              .MortPctType = kMortPctTypeSeriatimM1M2
                        End Select

                        .UseUnisex = (ParsedClassDefn(kRowsClasses.UseUnisex_GC)(classNum) = "Yes")
                        If .UseUnisex = True Then
                           .UniMalePct = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.UnisexMalePct_GC)(classNum))
                        Else
                           .UniMalePct = Nothing
                        End If

                        Select Case ParsedClassDefn(kRowsClasses.UsePreRetMort_GC)(classNum)
                           Case "Yes"
                              .UsePreRetMort = True
                           Case "No"
                              .UsePreRetMort = False
                           Case Else
                              .UsePreRetMort = Nothing
                        End Select

                        Select Case ParsedClassDefn(kRowsClasses.IntrRateOrVec_GC)(classNum)
                           Case "Rate"
                              .IntrMethod = Array.IndexOf(vIntrMethods, "Rate")
                              .IntrRate = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.IntrRate_GC)(classNum))
                           Case "Vector"
                              .IntrMethod = Array.IndexOf(vIntrMethods, "Vector")
                              InterpretVector(ParsedClassDefn(kRowsClasses.IntrVec_GC)(classNum), .IntrVec)
                              'tmpVec = DivaCalc.ParseVectorInput(ParsedClassDefn(kRowsClasses.IntrVec_GC)(classNum))
                              '.IntrVec = ExtendVector(tmpVec, VecMax)
                        End Select

                        Select Case ParsedClassDefn(kRowsClasses.InflRateOrVec_GC)(classNum)
                           Case "Rate"
                              .InflMethod = Array.IndexOf(vInflMethods, "Rate")
                              .InflRate = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.InflRate_GC)(classNum))
                           Case "Vector"
                              .InflMethod = Array.IndexOf(vInflMethods, "Vector")
                              InterpretVector(ParsedClassDefn(kRowsClasses.InflVec_GC)(classNum), .InflVec)
                              'tmpVec = DivaCalc.ParseVectorInput(ParsedClassDefn(kRowsClasses.InflVec_GC)(classNum))
                              '.InflVec = ExtendVector(tmpVec, VecMax)
                        End Select

                     Case 1   'SY Basis

                        .MortTableM = MortTableToIndex(ParsedClassDefn(kRowsClasses.MortTable1_SY)(classNum).Trim.ToUpper)
                        .MortTableF = MortTableToIndex(ParsedClassDefn(kRowsClasses.MortTable2_SY)(classNum).Trim.ToUpper)
                        .MortTableMYear = CInt(ParsedClassDefn(kRowsClasses.MortTableMYear_SY)(classNum))
                        .MortTableFYear = CInt(ParsedClassDefn(kRowsClasses.MortTableFYear_SY)(classNum))

                        If ParsedClassDefn(kRowsClasses.WxTable1_SY)(classNum).Trim = "" Then
                           .WxTableM = Nothing
                           .UseWithdrawalM = False
                        Else
                           .WxTableM = WxTableToIndex(ParsedClassDefn(kRowsClasses.WxTable1_SY)(classNum).Trim.ToUpper)
                           .UseWithdrawalM = True
                        End If

                        If ParsedClassDefn(kRowsClasses.WxTable2_SY)(classNum).Trim = "" Then
                           .WxTableF = Nothing
                           .UseWithdrawalF = False
                        Else
                           .WxTableF = WxTableToIndex(ParsedClassDefn(kRowsClasses.WxTable1_SY)(classNum).Trim.ToUpper)
                           .UseWithdrawalF = True
                        End If

                        If ParsedClassDefn(kRowsClasses.ImprovTable1_SY)(classNum).Trim = "" Then
                           .ImprovTableM = Nothing
                           .UseImprovementM = False
                           .ImprovFactorsM = Nothing
                        Else
                           .ImprovTableM = ImprovTableToIndex(ParsedClassDefn(kRowsClasses.ImprovTable1_SY)(classNum).Trim.ToUpper)
                           .UseImprovementM = True
                           .ImprovFactorsM = ComputeIxFactors(.ImprovTableM, .MortTableMYear, Math.Max(MortYearsToProject, 120))
                        End If

                        If ParsedClassDefn(kRowsClasses.ImprovTable2_SY)(classNum).Trim = "" Then
                           .ImprovTableF = Nothing
                           .UseImprovementF = False
                           .ImprovFactorsF = Nothing
                        Else
                           .ImprovTableF = ImprovTableToIndex(ParsedClassDefn(kRowsClasses.ImprovTable2_SY)(classNum).Trim.ToUpper)
                           .UseImprovementF = True
                           .ImprovFactorsF = ComputeIxFactors(.ImprovTableF, .MortTableFYear, Math.Max(MortYearsToProject, 120))
                        End If

                        'Retirement age must be rounded to nearest month in order to match up with ages in the retirement loop
                        .RetAge = Math.Round(DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.RetAge_SY)(classNum)) * 12.0) / 12.0

                        Select Case ParsedClassDefn(kRowsClasses.MortPctType_SY)(classNum)
                           Case "None"
                              .MortPctType = kMortPctTypeNone
                              .MortPctM = 1.0
                              .MortPctF = 1.0
                           Case "Inputs"
                              .MortPctType = kMortPctTypeInput
                              .MortPctM = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.MortPctMale_SY)(classNum))
                              .MortPctF = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.MortPctFem_SY)(classNum))
                           Case "Seriatim"
                              .MortPctType = kMortPctTypeSeriatim
                           Case "Seriatim [M1 * Qx + M2]"
                              .MortPctType = kMortPctTypeSeriatimM1M2
                        End Select

                        .UseUnisex = (ParsedClassDefn(kRowsClasses.UseUnisex_SY)(classNum) = "Yes")
                        If .UseUnisex = True Then
                           .UniMalePct = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.UnisexMalePct_SY)(classNum))
                        Else
                           .UniMalePct = Nothing
                        End If

                        Select Case ParsedClassDefn(kRowsClasses.UsePreRetMort_SY)(classNum)
                           Case "Yes"
                              .UsePreRetMort = True
                           Case "No"
                              .UsePreRetMort = False
                           Case Else
                              .UsePreRetMort = Nothing
                        End Select

                        Select Case ParsedClassDefn(kRowsClasses.IntrRateOrVec_SY)(classNum)
                           Case "Rate"
                              .IntrMethod = Array.IndexOf(vIntrMethods, "Rate")
                              .IntrRate = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.IntrRate_SY)(classNum))
                           Case "Vector"
                              .IntrMethod = Array.IndexOf(vIntrMethods, "Vector")
                              InterpretVector(ParsedClassDefn(kRowsClasses.IntrVec_SY)(classNum), .IntrVec)
                              'tmpVec = DivaCalc.ParseVectorInput(ParsedClassDefn(kRowsClasses.IntrVec_SY)(classNum))
                              '.IntrVec = ExtendVector(tmpVec, VecMax)
                        End Select

                        Select Case ParsedClassDefn(kRowsClasses.InflRateOrVec_SY)(classNum)
                           Case "Rate"
                              .InflMethod = Array.IndexOf(vInflMethods, "Rate")
                              .InflRate = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.InflRate_SY)(classNum))
                           Case "Vector"
                              .InflMethod = Array.IndexOf(vInflMethods, "Vector")
                              InterpretVector(ParsedClassDefn(kRowsClasses.InflVec_SY)(classNum), .InflVec)
                              'tmpVec = DivaCalc.ParseVectorInput(ParsedClassDefn(kRowsClasses.InflVec_SY)(classNum))
                              '.InflVec = ExtendVector(tmpVec, VecMax)
                        End Select

                     Case 2   'OT Basis
                        .MortTableM = MortTableToIndex(ParsedClassDefn(kRowsClasses.MortTable1_OT)(classNum).Trim.ToUpper)
                        .MortTableF = MortTableToIndex(ParsedClassDefn(kRowsClasses.MortTable2_OT)(classNum).Trim.ToUpper)
                        .MortTableMYear = CInt(ParsedClassDefn(kRowsClasses.MortTableMYear_OT)(classNum))
                        .MortTableFYear = CInt(ParsedClassDefn(kRowsClasses.MortTableFYear_OT)(classNum))

                        If ParsedClassDefn(kRowsClasses.WxTable1_OT)(classNum).Trim = "" Then
                           .WxTableM = Nothing
                           .UseWithdrawalM = False
                        Else
                           .WxTableM = WxTableToIndex(ParsedClassDefn(kRowsClasses.WxTable1_OT)(classNum).Trim.ToUpper)
                           .UseWithdrawalM = True
                        End If

                        If ParsedClassDefn(kRowsClasses.WxTable2_OT)(classNum).Trim = "" Then
                           .WxTableF = Nothing
                           .UseWithdrawalF = False
                        Else
                           .WxTableF = WxTableToIndex(ParsedClassDefn(kRowsClasses.WxTable1_OT)(classNum).Trim.ToUpper)
                           .UseWithdrawalF = True
                        End If

                        If ParsedClassDefn(kRowsClasses.ImprovTable1_OT)(classNum).Trim = "" Then
                           .ImprovTableM = Nothing
                           .UseImprovementM = False
                           .ImprovFactorsM = Nothing
                        Else
                           .ImprovTableM = ImprovTableToIndex(ParsedClassDefn(kRowsClasses.ImprovTable1_OT)(classNum).Trim.ToUpper)
                           .UseImprovementM = True
                           .ImprovFactorsM = ComputeIxFactors(.ImprovTableM, .MortTableMYear, Math.Max(MortYearsToProject, 120))
                        End If

                        If ParsedClassDefn(kRowsClasses.ImprovTable2_OT)(classNum).Trim = "" Then
                           .ImprovTableF = Nothing
                           .UseImprovementF = False
                           .ImprovFactorsF = Nothing
                        Else
                           .ImprovTableF = ImprovTableToIndex(ParsedClassDefn(kRowsClasses.ImprovTable2_OT)(classNum).Trim.ToUpper)
                           .UseImprovementF = True
                           .ImprovFactorsF = ComputeIxFactors(.ImprovTableF, .MortTableFYear, Math.Max(MortYearsToProject, 120))
                        End If

                        'Retirement age must be rounded to nearest month in order to match up with ages in the retirement loop
                        .RetAge = Math.Round(DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.RetAge_OT)(classNum)) * 12.0) / 12.0

                        Select Case ParsedClassDefn(kRowsClasses.MortPctType_OT)(classNum)
                           Case "None"
                              .MortPctType = kMortPctTypeNone
                           Case "Inputs"
                              .MortPctType = kMortPctTypeInput
                              .MortPctM = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.MortPctMale_OT)(classNum))
                              .MortPctF = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.MortPctFem_OT)(classNum))
                           Case "Seriatim"
                              .MortPctType = kMortPctTypeSeriatim
                           Case "Seriatim [M1 * Qx + M2]"
                              .MortPctType = kMortPctTypeSeriatimM1M2
                        End Select

                        .UseUnisex = (ParsedClassDefn(kRowsClasses.UseUnisex_OT)(classNum) = "Yes")
                        If .UseUnisex = True Then
                           .UniMalePct = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.UnisexMalePct_OT)(classNum))
                        Else
                           .UniMalePct = Nothing
                        End If

                        Select Case ParsedClassDefn(kRowsClasses.UsePreRetMort_OT)(classNum)
                           Case "Yes"
                              .UsePreRetMort = True
                           Case "No"
                              .UsePreRetMort = False
                           Case Else
                              .UsePreRetMort = Nothing
                        End Select

                        Select Case ParsedClassDefn(kRowsClasses.IntrRateOrVec_OT)(classNum)
                           Case "Rate"
                              .IntrMethod = Array.IndexOf(vIntrMethods, "Rate")
                              .IntrRate = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.IntrRate_OT)(classNum))
                           Case "Vector"
                              .IntrMethod = Array.IndexOf(vIntrMethods, "Vector")
                              InterpretVector(ParsedClassDefn(kRowsClasses.IntrVec_OT)(classNum), .IntrVec)
                              'tmpVec = DivaCalc.ParseVectorInput(ParsedClassDefn(kRowsClasses.IntrVec_OT)(classNum))
                              '.IntrVec = ExtendVector(tmpVec, VecMax)
                        End Select

                        Select Case ParsedClassDefn(kRowsClasses.InflRateOrVec_OT)(classNum)
                           Case "Rate"
                              .InflMethod = Array.IndexOf(vInflMethods, "Rate")
                              .InflRate = DivaCalc.EvaluateMathExpression(ParsedClassDefn(kRowsClasses.InflRate_OT)(classNum))
                           Case "Vector"
                              .InflMethod = Array.IndexOf(vInflMethods, "Vector")
                              InterpretVector(ParsedClassDefn(kRowsClasses.InflVec_OT)(classNum), .InflVec)
                              'tmpVec = DivaCalc.ParseVectorInput(ParsedClassDefn(kRowsClasses.InflVec_OT)(classNum))
                              '.InflVec = ExtendVector(tmpVec, VecMax)
                        End Select

                  End Select 'basis

                  If .IntrMethod = kMethodVector Then
                     Dim tmpFctr, tmpFctrStart, tmpFctrMid, tmpFctrEnd As Double
                     ReDim .IntrFctrStart(VecMax), .IntrFctrAnn(VecMax)

                     tmpFctrStart = 1.0
                     For k = 0 To VecMax
                        tmpFctr = (1.0 + .IntrVec(k)) ^ (-CalcExponent)
                        tmpFctrMid = tmpFctrStart * tmpFctr ^ 0.5
                        tmpFctrEnd = tmpFctrStart * tmpFctr
                        .IntrFctrStart(k) = tmpFctrStart
                        Select Case vCashFlowAnnuity
                           Case vCashFlowStart
                              .IntrFctrAnn(k) = tmpFctrStart
                           Case vCashFlowMiddle
                              .IntrFctrAnn(k) = tmpFctrMid
                           Case vCashFlowEnd
                              .IntrFctrAnn(k) = tmpFctrEnd
                        End Select
                        tmpFctrStart = tmpFctrEnd
                     Next
                  End If

                  ReDim .InflFctr(VecMax)
                  Select Case .InflMethod
                     Case kMethodVector
                        .InflFctr(0) = 1.0
                        For k = 0 To VecMax - 1
                           .InflFctr(k + 1) = .InflFctr(k) * ((1.0 + .InflVec(k)) ^ CalcExponent)
                        Next k
                     Case kMethodVector1PerYr
                        Dim tmpVal As Double
                        tmpVal = 1.0 + .InflVec(0)
                        For k = 0 To VecMax - CalcFreq Step CalcFreq
                           For k2 = k To k + CalcFreq - 1
                              .InflFctr(k2) = tmpVal
                           Next k2
                           tmpVal *= (1.0 + .InflVec(k + CalcFreq))
                        Next k
                        .InflFctr(VecMax) = tmpVal
                     Case kMethodSeriatim, kMethodRate, kMethodSeriatim1PerYr, kMethodRate1PerYr
                        Erase .InflVec
                  End Select

                  'ReDim IntrFctrStart(VecMax), IntrFctrAnn(VecMax)

                  'If running retirement loop, get age and factor array from classes tab
                  If UseRetireLoop Then
                     Try
                        Select Case basis
                           Case 0  'GC
                              .RetLoopAgeFactors = InterpolateRetireAgeFactors(InterpretRetireAgeFactors(ParsedClassDefn(kRowsClasses.RetLoopAgeFactors_GC)(classNum)))
                           Case 1  'SY
                              .RetLoopAgeFactors = InterpolateRetireAgeFactors(InterpretRetireAgeFactors(ParsedClassDefn(kRowsClasses.RetLoopAgeFactors_SY)(classNum)))
                           Case 2  'OT
                              .RetLoopAgeFactors = InterpolateRetireAgeFactors(InterpretRetireAgeFactors(ParsedClassDefn(kRowsClasses.RetLoopAgeFactors_OT)(classNum)))
                        End Select

                     Catch ex As Exception
                        Throw New Exception("Error in node " & vNodeName & ", Class " & classNum + 1 & _
                                            ": " & ex.Message)
                     End Try

                     'Now test to make sure the retirement loop contains the normal retirement age. First get the normal retirement age for this class:
                     Dim ClassNormalRetireAge As Double = ClassList(classNum, basis).RetAge 'CalcValues(kRowsParams.RetireAge)

                     'Now get the first and last ages to run (the InterpolateRetireAgeFactors subroutine called above ensures that the input ages are in order from lowest to highest)
                     Dim loopMinAge As Double = ClassList(classNum, basis).RetLoopAgeFactors(0, 0)
                     Dim loopMaxAgeInd As Integer = ClassList(classNum, basis).RetLoopAgeFactors.GetUpperBound(0)
                     Dim loopMaxAge As Double = ClassList(classNum, basis).RetLoopAgeFactors(loopMaxAgeInd, 0)

                     'If normal retirement age falls outside of the loop range, throw an error.
                     If ClassNormalRetireAge < loopMinAge Or ClassNormalRetireAge > loopMaxAge Then
                        Throw New Exception("Error in node " & vNodeName & ", Class " & classNum + 1 & _
                                            ": When running multiple retirement ages, the Normal Retirement Age must be within those ages.")
                     End If

                  End If

               End With
            Next basis
         Next classNum
         '''END INITIALIZE CLASS PARAMETERS


         '   Interpret Seriatim Data

         InterestRateCollection = New Collections.Specialized.StringCollection
         InflationRateCollection = New Collections.Specialized.StringCollection

         SeriaMax = splParam(sTabSeriatimIn).GetUpperBound(0)
         If IsNothing(Seriatim) OrElse Seriatim.GetUpperBound(0) <> SeriaMax Then
            ReDim Seriatim(SeriaMax, 2)   'Second dimension is basis
            ReDim SeriatimOut(SeriaMax + 1)
         End If
         If gv.HasTimeSeries Then
            ValDate = gv.TimeSeries(0)
         Else
            FireMsgBox("In " & vNodeName & ": " & "Model must be run with time series enabled. Change the setting in the Parameters tab.", MsgBoxStyle.Critical)
         End If

         If HasAnySeriatimOut Or HasAnySupOut Then
            For k = 0 To SumOuts.GetUpperBound(0)
               SumOuts(k) = 0.0
            Next k
         End If

         WarnCount = 0
         For k = 0 To SeriaMax
            For basis = 0 To 2
               splLine = Split(splParam(sTabSeriatimIn)(k), vbTab)
               With Seriatim(k, basis)
                  Try
                     .Input = splLine
                     .classNum = classNames.Item(splLine(kColsSeriatimIn.className))

                     If UseAge Then
                        .Age1 = CDbl(splLine(kColsSeriatimIn.AgeDOB1))
                     Else
                        .DOB1 = GetSeriatimDate(splLine(kColsSeriatimIn.AgeDOB1))
                        If DateType = 3 Then
                           '   Year only
                           .Age1 = ValDate.Year - .DOB1.Year
                        Else
                           .Age1 = DateDiff(DateInterval.Month, .DOB1, ValDate) / 12.0
                           If .DOB1.Day > ValDate.Day Then .Age1 -= OneTwelfth
                        End If
                     End If

                     Select Case splLine(kColsSeriatimIn.Gender1).Trim
                        Case "M", "H", ""
                           .MortTable1 = ClassList(.classNum - 1, basis).MortTableM
                           .ImprovFactors1 = ClassList(.classNum - 1, basis).ImprovFactorsM
                           .WxTable1 = ClassList(.classNum - 1, basis).WxTableM
                           .IsMale = True
                        Case Else
                           .MortTable1 = ClassList(.classNum - 1, basis).MortTableF
                           .ImprovFactors1 = ClassList(.classNum - 1, basis).ImprovFactorsF
                           .WxTable1 = ClassList(.classNum - 1, basis).WxTableF
                           .IsMale = False
                     End Select

                     Select Case splLine(kColsSeriatimIn.JS).Trim
                        Case "", "S"
                           .JointIndex = kJointIsSingle
                        Case "J"
                           .JointIndex = kJointReduceEither
                        Case "J1"
                           .JointIndex = kJointReduceLife1
                        Case "J2"
                           .JointIndex = kJointReduceLife2
                        Case Else
                           .JointIndex = kJointIsSingle
                           WarnCount += 1
                           If WarnCount <= 10 Then
                              FireMsgBox("In Node '" & vNodeName & "', unknown joint status encountered: '" & splLine(kColsSeriatimIn.JS).Trim & "'" _
                                                     & vbCrLf & "Allowable choices are blank, 'S', 'J', 'J1', 'J2'.  Single Life Assumed.")
                           End If
                           If WarnCount = 10 Then
                              FireMsgBox(vbCrLf & "In Node '" & vNodeName & "', 10 warnings reached.  Further messages suppressed.")
                           End If
                     End Select
                     If .JointIndex >= 0 Then
                        'If UseSpouseAgeOffset Then
                        '	.Age2 = .Age1 + YrsSpouseAgeOffset
                        If UseAge Then 'Used to be Elseif 07/19/2016
                           .Age2 = CDbl(splLine(kColsSeriatimIn.AgeDOB2))
                        Else
                           .DOB2 = GetSeriatimDate(splLine(kColsSeriatimIn.AgeDOB2))
                           If DateType = 3 Then
                              '   Year only
                              .Age2 = ValDate.Year - .DOB2.Year
                           Else
                              .Age2 = DateDiff(DateInterval.Month, .DOB2, ValDate) / 12.0
                              If .DOB2.Day > ValDate.Day Then .Age2 -= OneTwelfth
                           End If
                        End If


                        Select Case splLine(kColsSeriatimIn.Gender2).Trim
                           Case "M", "H", ""
                              .MortTable2 = ClassList(.classNum - 1, basis).MortTableM
                              .ImprovFactors2 = ClassList(.classNum - 1, basis).ImprovFactorsM
                              .WxTable2 = ClassList(.classNum - 1, basis).WxTableM
                              .IsMale2 = True
                           Case Else
                              .MortTable2 = ClassList(.classNum - 1, basis).MortTableF
                              .ImprovFactors2 = ClassList(.classNum - 1, basis).ImprovFactorsF
                              .WxTable2 = ClassList(.classNum - 1, basis).WxTableF
                              .IsMale2 = False
                        End Select

                        'If splLine(kColsSeriatimIn.Smoking2).Trim = "S" OrElse splLine(kColsSeriatimIn.Smoking2).Trim = "F" Then .Table2 += 1
                        .SurvivorPct = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.SurvPct)) '8))
                     End If
                     '.InsBenefit = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.InsuBenefit))
                     '.AnnBenefit = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.AnnuBenefit)) * CalcExponent * AnnBenFreq

                     'If RetireUseDate Then
                     '    .RetireDate = GetSeriatimDate(splLine(11))
                     'End If
                     'Select Case RetireAgeDate
                     '	Case kRetireAgeParameter
                     '		'   Nothing
                     '	Case kRetireDateSeriatim
                     '		.RetireDate = GetSeriatimDate(splLine(kColsSeriatimIn.RetDateAge))
                     '	Case kRetireAgeSeriatim
                     '		.RetireAgeSeriatim = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.RetDateAge))
                     'End Select

                     ' keyINSUREMOVE
                     'If UsePostRetireIns Then
                     '	.InsBenefitPostRetire = 0 'DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.PostRetInsu))
                     'Else
                     '	.InsBenefitPostRetire = .InsBenefit
                     'End If
                     'If UsePostRetireAnn Then
                     .AnnBenefit = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.PostRetAnn)) * CalcExponent * AnnBenFreq
                     'Else
                     '              .AnnBenefitPostRetire = .AnnBenefit
                     'End If
                     Select Case LastGteeType
                        Case kLastGteeNone
                           .LastGteeDate = DateSerial(1900, 1, 1)
                           .LastGteeYears = 0.0
                        Case kLastGteeDate
                           .LastGteeDate = GetSeriatimDate(splLine(kColsSeriatimIn.LastGteeDate))
                           .LastGteeYears = 0.0
                        Case kLastGteeYears
                           .LastGteeYears = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.LastGteeDate))
                        Case kLastGteeFixed
                           .LastGteeDate = DateSerial(1900, 1, 1)
                     End Select
                     'If LastGteeUseDate Then
                     '    .LastGteeDate = GetSeriatimDate(splLine(14))
                     'Else
                     '    .LastGteeDate = DateSerial(1900, 1, 1)
                     'End If
                     'Select Case IntrMethod
                     '	Case kMethodSeriatim, kMethodSeriatimVector
                     '		tmpIntRateStr = splLine(kColsSeriatimIn.IntrRate).Trim.ToUpper
                     '		If InterestRateCollection.Contains(tmpIntRateStr) Then
                     '			.InterestIndex = InterestRateCollection.IndexOf(tmpIntRateStr)
                     '		Else
                     '			.InterestIndex = InterestRateCollection.Add(tmpIntRateStr)
                     '		End If
                     '	Case Else
                     '		.InterestIndex = 0
                     'End Select
                     If vMortPctType = kMortPctTypeSeriatim Then
                        .MortPct1 = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.MortPct1))
                        .MortPct2 = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.MortPct2))
                     ElseIf vMortPctType = kMortPctTypeSeriatimM1M2 Then
                        .MortPct1 = DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.MortPct1))
                        .MortPct2 = 0.001 * DivaCalc.EvaluateMathExpression(splLine(kColsSeriatimIn.MortPct2))
                     Else
                        .MortPct1 = 1.0
                        .MortPct2 = 1.0
                     End If
                     'tmpStr = splLine(kColsSeriatimIn.MortImprvScale).Trim.PadRight(2)
                     'Select Case tmpStr.Substring(0, 1)
                     '   Case "B"
                     '      .MortImprScale1 = kScaleB
                     '   Case Else
                     '      .MortImprScale1 = kScaleA
                     'End Select
                     'Select Case tmpStr.Substring(1, 1)
                     '   Case "B"
                     '      .MortImprScale2 = kScaleB
                     '   Case Else
                     '      .MortImprScale2 = kScaleA
                     'End Select
                     '.AgeBandIndex = MaxAgeBandList
                     'For k2 = MaxAgeBandList To 0 Step -1
                     '	If AgeBandList(k2) <= .Age1 Then
                     '		.AgeBandIndex = k2
                     '		Exit For
                     '	End If
                     'Next k2

                  Catch ex As Exception
                     WarnCount += 1
                     If WarnCount <= 10 Then
                        FireMsgBox("In " & vNodeName & ", Seriatim data line " & CStr(k) & ", Name " & .Input(0) & ":" & vbCrLf & ex.Message)
                     End If
                     If WarnCount = 10 Then
                        FireMsgBox(vbCrLf & "In Node '" & vNodeName & "', 10 warnings reached.  Further messages suppressed.")
                     End If
                  End Try
               End With
            Next basis
         Next k

         '   Create arrays to hold Supplementary detailed output
         If HasAnySupOut AndAlso TimePeriodsToDump >= 0 Then
            ReDim SuppOut(TimePeriodsToDump)
            For k = 0 To TimePeriodsToDump
               ReDim SuppOut(k)(SeriaMax)
            Next k
         Else
            Erase SuppOut
         End If
         Return True

      Catch ex As Exception
         FireMsgBox("In " & vNodeName & ": " & ex.Message, MsgBoxStyle.Critical)
         Return False
      End Try
   End Function

   Public Overrides Sub DoCalcs(ByVal TimePeriod As Integer)
      '    This section performs all calculations
      Dim k, tp, kItem, basis, kDur, DurYr, PrevDurYr, VecMaxThis1, VecMaxThis2 As Integer
      Dim tpSeria, tpToUse As Integer
      Dim tmpVals(10), prevLivesInd As Double
      Dim TblAgeX, TblAge1, TblAge2 As Integer, tblAgeFractional, AgeInterp1, AgeInterp2 As Double
      Dim tmpFctr, MortPct, MortAdd As Double
      Dim NormalRetireAge As Double
      Dim RetireDur, GteeAnnDur, MortStartDur As Integer
      Dim tmpQx, tmpQxM, tmpQxF, tmpWx, tmpLxZ, tmpDeltaLx As Double, tmpAge As Integer
      Dim incrImp, incrImp1, incrImp2 As Double 'Mortality improvement to apply to trend mortality to the day within the year. Vars with 1 and 2 are for interpolate age settings.
      Dim tmpQx1, tmpQx2, tmpQx1M, tmpQx1F, tmpQx2M, tmpQx2F, tmpWx1, tmpWx2 As Double
      Dim AnnSumPost, RetireAnn As Double
      Dim tmpOutVal As Double
      Dim SumAnn, SumAnnPost, NumTimeItems As Double
      Dim intNumTimeItems As Integer
      Dim SumCF, SumBestCF As Double
      Dim SumAnnPv As Double
      Dim tpS, tpE, tX As Integer
      Dim NeedTimeSum As Boolean
      Dim tmpQxTableM, tmpQxTableF As MortTable
      Dim tmpIxTableM, tmpIxTableF As ImprovementFactors

      Dim calYear As Integer, fracYear As Double
      'calYear is the calendar year at each duration in DoCalcs, fracYear is the number of months into the year at each duration in DoCalcs
      '(Used for mortality improvement.)

      Dim EarliestRetireAge As Double
      Dim EarlyRetireDur As Integer

      'Dim testYear As DateTime
      'Dim YrsMortImpGrade As Double

      MortAdd = 0.0 '   Only redefined in one mortality pct type
      tmpLxZ = Double.NaN    '   Set these to cause errors in case of programming hiccup
      tmpDeltaLx = Double.NaN

      '   This whole routine works only on time period 1 values then projects out

      If TimePeriod = 1 Then
         '   Zero-out age-bucket accumulators
         NeedTimeSum = (CalcFreq <> gv.TimePeriodsPerYear)
         NumTimeItems = CalcFreq / gv.TimePeriodsPerYear
         intNumTimeItems = CInt(NumTimeItems)

         'Added 07/20/2016
         CalcValues(0) = New Double(sMaxParams) {}
         CalcValues(1) = New Double(kRowsClasses.LastRow - 1) {}

         For Each k In vRowsToRecalc(sTabParameters)
            Select Case CellType(sTabParameters)(k, 1)
               Case DivaCalcTools.GridConstants.sCellTypeFixed
                  CalcValues(sTabParameters)(k) = ParmValues(sTabParameters)(k, 1)
               Case DivaCalcTools.GridConstants.sCellTypeFormula
                  CalcValues(sTabParameters)(k) = DivaCalc.NodeEvalueSequence(FuncExpressions(sTabParameters)(k, 1), TimePeriod, vEleNames, vEleVal, vNodeName)
               Case DivaCalcTools.GridConstants.sCellTypeNode
                  CalcValues(sTabParameters)(k) = ParmNode(sTabParameters)(k, 1).ValueElementTime(ParmElement(sTabParameters)(k, 1), TimePeriod)
            End Select
         Next k

         'RECALC CLASSES
         For kClass As Integer = 0 To CellType(sTabClasses).GetUpperBound(1) - 1
            For Each k In vRowsToRecalc(sTabClasses)
               Select Case CellType(sTabClasses)(k, kClass + 1)
                  Case DivaCalcTools.GridConstants.sCellTypeFormula
                     ParmValues(sTabClasses)(k, kClass + 1) = DivaCalc.NodeEvalueSequence(FuncExpressions(sTabClasses)(k, kClass + 1), TimePeriod, vEleNames, vEleVal, vNodeName)
                  Case DivaCalcTools.GridConstants.sCellTypeNode
                     ParmValues(sTabClasses)(k, kClass + 1) = ParmNode(sTabClasses)(k, 1).ValueElementTime(ParmElement(sTabClasses)(k, kClass + 1), TimePeriod)
               End Select
            Next k

            If ClassList(kClass, 0).IntrMethod = kMethodRate Then
               ClassList(kClass, 0).IntrRate = ParmValues(sTabClasses)(kRowsClasses.IntrRate_GC, kClass + 1)
               ClassList(kClass, 0).InflRate = ParmValues(sTabClasses)(kRowsClasses.InflRate_GC, kClass + 1)
            End If

            If ClassList(kClass, 1).IntrMethod = kMethodRate Then
               ClassList(kClass, 1).IntrRate = ParmValues(sTabClasses)(kRowsClasses.IntrRate_SY, kClass + 1)
               ClassList(kClass, 1).InflRate = ParmValues(sTabClasses)(kRowsClasses.InflRate_SY, kClass + 1)
            End If

            If ClassList(kClass, 2).IntrMethod = kMethodRate Then
               ClassList(kClass, 2).IntrRate = ParmValues(sTabClasses)(kRowsClasses.IntrRate_OT, kClass + 1)
               ClassList(kClass, 2).InflRate = ParmValues(sTabClasses)(kRowsClasses.InflRate_OT, kClass + 1)
            End If

         Next kClass

         'END RECALC CLASSES

         For k = 0 To VecMax

            AnnPmtSumPost(k) = 0.0

            For kBasis As Integer = 0 To 2
               LivesSum(kBasis, k) = 0.0
            Next kBasis

            npvAnnSumPost(k) = 0.0
            npvAnnSumNormalAge(k) = 0.0
         Next k

         LivesInd(0) = 1.0
         LivesInd2(0) = 1.0

         For kItem = 0 To SeriaMax
            For basis = 0 To 2
               Try
                  With Seriatim(kItem, basis)

                     If ClassList(.classNum - 1, basis).UseUnisex = True Then
                        UniMalePct = ClassList(.classNum - 1, basis).UniMalePct
                        UniFemPct = 1.0 - UniMalePct
                     Else
                        UniMalePct = Nothing
                        UniFemPct = Nothing
                     End If

                     tmpQxTableM = ClassList(.classNum - 1, basis).MortTableM
                     tmpQxTableF = ClassList(.classNum - 1, basis).MortTableF
                     tmpIxTableM = ClassList(.classNum - 1, basis).ImprovFactorsM
                     tmpIxTableF = ClassList(.classNum - 1, basis).ImprovFactorsF

                     NormalRetireAge = ClassList(.classNum - 1, basis).RetAge 'CalcValues(kRowsParams.RetireAge)

                     'Calculate interest factors and store in vector:
                     Select Case ClassList(.classNum - 1, basis).IntrMethod
                        Case kMethodRate
                           ReDim ClassList(.classNum - 1, basis).IntrFctrStart(VecMax), ClassList(.classNum - 1, basis).IntrFctrAnn(VecMax)
                           tmpFctr = (1.0 + ClassList(.classNum - 1, basis).IntrRate) ^ (-CalcExponent)
                           ClassList(.classNum - 1, basis).IntrFctrStart(0) = 1.0

                           Select Case vCashFlowAnnuity
                              Case vCashFlowStart
                                 ClassList(.classNum - 1, basis).IntrFctrAnn(0) = 1.0
                              Case vCashFlowMiddle
                                 ClassList(.classNum - 1, basis).IntrFctrAnn(0) = tmpFctr ^ 0.5
                              Case vCashFlowEnd
                                 ClassList(.classNum - 1, basis).IntrFctrAnn(0) = tmpFctr
                           End Select
                           Dim kMinus1 As Integer
                           For k = 1 To VecMax
                              kMinus1 = k - 1
                              ClassList(.classNum - 1, basis).IntrFctrStart(k) = ClassList(.classNum - 1, basis).IntrFctrStart(kMinus1) * tmpFctr
                              ClassList(.classNum - 1, basis).IntrFctrAnn(k) = ClassList(.classNum - 1, basis).IntrFctrAnn(kMinus1) * tmpFctr
                           Next k

                        Case kMethodVector
                           '   Nothing - already handled
                     End Select

                     Select Case ClassList(.classNum - 1, basis).InflMethod
                        Case kMethodRate
                           ClassList(.classNum - 1, basis).InflFctr(0) = 1.0
                           tmpFctr = (1.0 + ClassList(.classNum - 1, basis).InflRate) ^ CalcExponent
                           For k = 1 To VecMax
                              ClassList(.classNum - 1, basis).InflFctr(k) = ClassList(.classNum - 1, basis).InflFctr(k - 1) * tmpFctr
                           Next k

                        Case kMethodRate1PerYr
                           tmpFctr = (1.0 + ClassList(.classNum - 1, basis).InflRate)
                           ClassList(.classNum - 1, basis).InflFctr(0) = tmpFctr
                           For k = 1 To VecMax
                              If k Mod CalcFreq = 0 Then
                                 ClassList(.classNum - 1, basis).InflFctr(k) = ClassList(.classNum - 1, basis).InflFctr(k - 1) * tmpFctr
                              Else
                                 ClassList(.classNum - 1, basis).InflFctr(k) = ClassList(.classNum - 1, basis).InflFctr(k - 1)
                              End If
                           Next k

                        Case kMethodVector, kMethodVector1PerYr
                           '   Already handled
                     End Select

                     VecMaxThis1 = VecMax
                     VecMaxThis2 = VecMax

                     Select Case ClassList(.classNum - 1, basis).MortPctType
                        Case kMortPctTypeNone
                           MortPct = 1.0
                        Case kMortPctTypeInput
                           If .IsMale Then
                              MortPct = ClassList(.classNum - 1, basis).MortPctM
                           Else
                              MortPct = ClassList(.classNum - 1, basis).MortPctF
                           End If
                        Case kMortPctTypeSeriatim
                           MortPct = .MortPct1
                        Case kMortPctTypeSeriatimM1M2
                           MortPct = .MortPct1
                           MortAdd = .MortPct2
                     End Select

                     tblAgeFractional = .Age1

                     Select Case vAgeBasis
                        Case AgeBasisLast
                           TblAgeX = CInt(Math.Floor(tblAgeFractional))
                        Case AgeBasisNearest
                           TblAgeX = CInt(Math.Round(tblAgeFractional))
                        Case AgeBasisInterpolate
                           TblAge1 = CInt(Math.Floor(tblAgeFractional))
                           TblAge2 = 1 + TblAge1
                           AgeInterp2 = tblAgeFractional - TblAge1
                           AgeInterp1 = 1.0 - AgeInterp2
                        Case AgeBasisInterpolateMonths
                           TblAge1 = CInt(Math.Floor(tblAgeFractional))
                           TblAge2 = 1 + TblAge1
                           AgeInterp2 = Math.Floor(12.0 * (tblAgeFractional - TblAge1)) / 12.0
                           AgeInterp1 = 1.0 - AgeInterp2
                     End Select


                     .RetireAgeSeriatim = NormalRetireAge
                     RetireDur = CInt((NormalRetireAge - .Age1) * CalcFreq)

                     'If retirement loop is on, and if pre-retirement mortality is not used, then mortality needs to start at age of earliest retirement
                     'Then the actual mortality figures can be backed out by dividing by the survival probability to each retirement age used in the calculation
                     'This isn't a problem if the retirement loop is off
                     'Even if loop is on, this isn't a problem if pre-retirement mortality is used, because then the retirement duration doesn't affect the survival probabilities
                     If UseRetireLoop And Not UseMortPreRetire Then
                        EarliestRetireAge = ClassList(.classNum - 1, basis).RetLoopAgeFactors(0, 0)
                        EarlyRetireDur = CInt((EarliestRetireAge - .Age1) * CalcFreq)
                        RetireDur = Math.Min(RetireDur, EarlyRetireDur)
                     End If

                     'Don't think this next bit of code is needed to guarantee is done within retirement loop now. 
                     'Kept it in for now but look into deleting it later.
                     Select Case LastGteeType
                        Case kLastGteeNone
                           GteeAnnDur = 0
                        Case kLastGteeDate
                           GteeAnnDur = CInt(DateDiff(DateInterval.Month, ValDate, .LastGteeDate) / 12.0 * CalcFreq)
                        Case kLastGteeYears   '   Seriatim years are from valuation date
                           GteeAnnDur = Math.Max(0, CInt(RetireDur + .LastGteeYears * CalcFreq))
                        Case kLastGteeFixed   '   Preset years are from retirement date
                           GteeAnnDur = Math.Max(0, CInt(RetireDur + YrsGteeAnn * CalcFreq))
                     End Select

                     If RetireDur < 0 Then RetireDur = 0

                     If ClassList(.classNum - 1, basis).UsePreRetMort Then
                        MortStartDur = 0
                        LivesInd(0) = 1.0
                        LivesInd2(0) = 1.0
                     Else
                        MortStartDur = RetireDur
                        For kDur = 0 To MortStartDur
                           LivesInd(kDur) = 1.0
                           LivesInd2(kDur) = 1.0
                        Next kDur
                     End If
                     PrevDurYr = -1

                     Select Case vAgeBasis
                        Case AgeBasisLast, AgeBasisNearest
                           For kDur = MortStartDur To VecMax - 1
                              'Get the date at this duration, along with the day of year, fraction through year, and calendar year
                              'These are used for mortality improvement.
                              currDate = ValDate.AddMonths(kDur * (12 \ CalcFreq))
                              currDay = currDate.DayOfYear
                              fracYear = currDay / 365
                              calYear = currDate.Year
                              'calYear = ValDate.AddMonths((kDur + 1) * (12 \ CalcFreq)).Year
                              DurYr = CInt(kDur \ CalcFreq)

                              If DurYr <> PrevDurYr Then
                                 PrevDurYr = DurYr

                                 If (.IsMale And ClassList(.classNum - 1, basis).UseWithdrawalM) Or ((Not .IsMale) And ClassList(.classNum - 1, basis).UseWithdrawalF) Then
                                    tmpWx = .WxTable1.Wx(TblAgeX)
                                 Else
                                    tmpWx = 0
                                 End If

                                 If DurYr <= .MortTable1.SelectPer Then 'Still in select period

                                    If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale Then
                                       tmpQxM = tmpQxTableM.Qx(TblAgeX, DurYr)
                                       If ClassList(.classNum - 1, basis).UseImprovementM Then
                                          'We need to trend to the exact point in the year
                                          'incrImp is the partial year of improvement needed for the current year
                                          'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                          'Then raise this to the fraction through the year at the current date
                                          incrImp = (tmpIxTableM.Ix(TblAgeX, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(TblAgeX, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                          tmpQxM *= tmpIxTableM.Ix(TblAgeX, calYear - tmpIxTableM.BaseYear) * incrImp
                                       End If
                                    End If

                                    If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale Then
                                       tmpQxF = tmpQxTableF.Qx(TblAgeX, DurYr)
                                       If ClassList(.classNum - 1, basis).UseImprovementF Then
                                          'We need to trend to the exact point in the year
                                          'incrImp is the partial year of improvement needed for the current year
                                          'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                          'Then raise this to the fraction through the year at the current date
                                          incrImp = (tmpIxTableF.Ix(TblAgeX, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(TblAgeX, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                          tmpQxF *= tmpIxTableF.Ix(TblAgeX, calYear - tmpIxTableF.BaseYear) * incrImp
                                       End If
                                    End If

                                    If ClassList(.classNum - 1, basis).UseUnisex Then
                                       tmpQx = tmpQxM * UniMalePct + tmpQxF * UniFemPct
                                    ElseIf .IsMale Then
                                       tmpQx = tmpQxM
                                    ElseIf Not .IsMale Then
                                       tmpQx = tmpQxF
                                    Else
                                       FireMsgBox("ERROR CODE A")
                                    End If

                                 Else 'If not in select period...
                                    tmpAge = TblAgeX + DurYr - .MortTable1.SelectPer 'Define effective age for ultimate lookup. E.g. IssAge 30, Duration 15, Select Period 10, then effective age is 35 for ultimate period lookup
                                    If tmpAge <= .MortTable1.EndAge Then
                                       'Note that even if there is technically an entry for tmpAge, this won't work. E.g. if table "ends" at 35, select period is 10 years,
                                       'then technically you can look up an ultimate age 45 rate. But this logic will say that the table ends at 35, so won't look it up.
                                       If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale Then
                                          tmpQxM = tmpQxTableM.Qx(tmpAge, tmpQxTableM.SelectPer)
                                          If ClassList(.classNum - 1, basis).UseImprovementM Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp = (tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                             tmpQxM *= tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear) * incrImp
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale Then
                                          tmpQxF = tmpQxTableF.Qx(tmpAge, tmpQxTableF.SelectPer)
                                          If ClassList(.classNum - 1, basis).UseImprovementF Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp = (tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                             tmpQxF *= tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear) * incrImp
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Then
                                          tmpQx = tmpQxM * UniMalePct + tmpQxF * UniFemPct
                                       ElseIf .IsMale Then
                                          tmpQx = tmpQxM
                                       ElseIf Not .IsMale Then
                                          tmpQx = tmpQxF
                                       Else
                                          FireMsgBox("ERROR CODE A")
                                       End If

                                    Else 'If effective age is greater than the max table age...
                                       tmpQx = 0.0 'Why set to zero? Why not 1?
                                       VecMaxThis1 = kDur
                                       LivesInd(kDur + 1) = LivesInd(kDur) * tmpLxZ
                                       Exit For 'Is this on error? Does this interact ok with withdrawal rates?
                                    End If
                                 End If
                                 If vIntraYearIsExponential Then
                                    tmpLxZ = (1.0 - Math.Max(Math.Min(MortAdd + tmpQx * MortPct + tmpWx, 1), 0)) ^ CalcExponent
                                 Else
                                    tmpDeltaLx = LivesInd(kDur) * (MortAdd + tmpQx * MortPct + tmpWx) * CalcExponent
                                 End If
                              End If

                              If vIntraYearIsExponential Then
                                 LivesInd(kDur + 1) = LivesInd(kDur) * tmpLxZ
                              Else
                                 LivesInd(kDur + 1) = Math.Max(0.0, LivesInd(kDur) - tmpDeltaLx)
                              End If
                           Next kDur

                        Case AgeBasisInterpolate, AgeBasisInterpolateMonths
                           For kDur = MortStartDur To VecMax - 1
                              'Get the date at this duration, along with the day of year, fraction through year, and calendar year
                              'These are used for mortality improvement.
                              currDate = ValDate.AddMonths(kDur * (12 \ CalcFreq))
                              currDay = currDate.DayOfYear
                              fracYear = currDay / 365
                              calYear = currDate.Year

                              'calYear = ValDate.AddMonths((kDur + 1) * (12 \ CalcFreq)).Year
                              If vAgeBasis = AgeBasisInterpolateMonths Then
                                 DurYr = CInt(Int((kDur + AgeInterp2 * 12) / CalcFreq))
                              Else
                                 DurYr = CInt(Int(kDur / CalcFreq))
                              End If
                              If DurYr <> PrevDurYr Then
                                 PrevDurYr = DurYr
                                 If (.IsMale And ClassList(.classNum - 1, basis).UseWithdrawalM) Or ((Not .IsMale) And ClassList(.classNum - 1, basis).UseWithdrawalF) Then
                                    tmpWx1 = .WxTable1.Wx(TblAge1)
                                    tmpWx2 = .WxTable1.Wx(TblAge2)
                                 Else
                                    tmpWx1 = 0
                                    tmpWx2 = 0
                                 End If

                                 If DurYr <= .MortTable1.SelectPer Then
                                    If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale Then
                                       tmpQx1M = tmpQxTableM.Qx(TblAge1, DurYr)
                                       tmpQx2M = tmpQxTableM.Qx(TblAge2, DurYr)
                                       If ClassList(.classNum - 1, basis).UseImprovementM Then
                                          'We need to trend to the exact point in the year
                                          'incrImp is the partial year of improvement needed for the current year
                                          'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                          'Then raise this to the fraction through the year at the current date
                                          incrImp1 = (tmpIxTableM.Ix(TblAge1, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(TblAge1, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                          incrImp2 = (tmpIxTableM.Ix(TblAge2, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(TblAge2, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                          tmpQx1M *= tmpIxTableM.Ix(TblAge1, calYear - tmpIxTableM.BaseYear) * incrImp1
                                          tmpQx2M *= tmpIxTableM.Ix(TblAge2, calYear - tmpIxTableM.BaseYear) * incrImp2
                                       End If
                                    End If

                                    If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale Then
                                       tmpQx1F = tmpQxTableF.Qx(TblAge1, DurYr)
                                       tmpQx2F = tmpQxTableF.Qx(TblAge2, DurYr)
                                       If ClassList(.classNum - 1, basis).UseImprovementF Then
                                          'We need to trend to the exact point in the year
                                          'incrImp is the partial year of improvement needed for the current year
                                          'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                          'Then raise this to the fraction through the year at the current date
                                          incrImp1 = (tmpIxTableF.Ix(TblAge1, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(TblAge1, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                          incrImp2 = (tmpIxTableF.Ix(TblAge2, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(TblAge2, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                          tmpQx1F *= tmpIxTableF.Ix(TblAge1, calYear - tmpIxTableF.BaseYear) * incrImp1
                                          tmpQx2F *= tmpIxTableF.Ix(TblAge2, calYear - tmpIxTableF.BaseYear) * incrImp2
                                       End If
                                    End If

                                    If ClassList(.classNum - 1, basis).UseUnisex Then
                                       tmpQx1 = tmpQx1M * UniMalePct + tmpQx1F * UniFemPct
                                       tmpQx2 = tmpQx2M * UniMalePct + tmpQx2F * UniFemPct
                                    ElseIf .IsMale Then
                                       tmpQx1 = tmpQx1M
                                       tmpQx2 = tmpQx2M
                                    ElseIf Not .IsMale Then
                                       tmpQx1 = tmpQx1F
                                       tmpQx2 = tmpQx2F
                                    Else
                                       FireMsgBox("ERROR CODE A")
                                    End If

                                 Else
                                    tmpAge = TblAge1 + DurYr - .MortTable1.SelectPer

                                    If tmpAge < .MortTable1.EndAge Then   '   Note use of < rather than <= due to interpolation
                                       If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale Then
                                          tmpQx1M = tmpQxTableM.Qx(tmpAge, tmpQxTableM.SelectPer)
                                          tmpQx2M = tmpQxTableM.Qx(tmpAge + 1, tmpQxTableM.SelectPer)
                                          If ClassList(.classNum - 1, basis).UseImprovementM Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp1 = (tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                             incrImp2 = (tmpIxTableM.Ix(tmpAge + 1, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(tmpAge + 1, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                             tmpQx1M *= tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear) * incrImp1
                                             tmpQx2M *= tmpIxTableM.Ix(tmpAge + 1, calYear - tmpIxTableM.BaseYear) * incrImp2
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale Then
                                          tmpQx1F = tmpQxTableF.Qx(tmpAge, tmpQxTableF.SelectPer)
                                          tmpQx2F = tmpQxTableF.Qx(tmpAge + 1, tmpQxTableF.SelectPer)
                                          If ClassList(.classNum - 1, basis).UseImprovementF Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp1 = (tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                             incrImp2 = (tmpIxTableF.Ix(tmpAge + 1, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(tmpAge + 1, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                             tmpQx1F *= tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear) * incrImp1
                                             tmpQx2F *= tmpIxTableF.Ix(tmpAge + 1, calYear - tmpIxTableF.BaseYear) * incrImp2
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Then
                                          tmpQx1 = tmpQx1M * UniMalePct + tmpQx1F * UniFemPct
                                          tmpQx2 = tmpQx2M * UniMalePct + tmpQx2F * UniFemPct
                                       ElseIf .IsMale Then
                                          tmpQx1 = tmpQx1M
                                          tmpQx2 = tmpQx2M
                                       ElseIf Not .IsMale Then
                                          tmpQx1 = tmpQx1F
                                          tmpQx2 = tmpQx2F
                                       Else
                                          FireMsgBox("ERROR CODE A")
                                       End If

                                    ElseIf tmpAge = .MortTable1.EndAge Then     '   Assume last Qx is 1.00

                                       If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale Then
                                          tmpQx1M = tmpQxTableM.Qx(tmpAge, tmpQxTableM.SelectPer)
                                          tmpQx2M = 1
                                          If ClassList(.classNum - 1, basis).UseImprovementM Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp1 = (tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                             tmpQx1M *= tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear) * incrImp1
                                             tmpQx2M *= 1
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale Then
                                          tmpQx1F = tmpQxTableF.Qx(tmpAge, tmpQxTableF.SelectPer)
                                          tmpQx2F = 1
                                          If ClassList(.classNum - 1, basis).UseImprovementF Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp1 = (tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                             tmpQx1F *= tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear) * incrImp1
                                             tmpQx2F *= 1
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Then
                                          tmpQx1 = tmpQx1M * UniMalePct + tmpQx1F * UniFemPct
                                          tmpQx2 = tmpQx2M * UniMalePct + tmpQx2F * UniFemPct
                                       ElseIf .IsMale Then
                                          tmpQx1 = tmpQx1M
                                          tmpQx2 = tmpQx2M
                                       ElseIf Not .IsMale Then
                                          tmpQx1 = tmpQx1F
                                          tmpQx2 = tmpQx2F
                                       Else
                                          FireMsgBox("ERROR CODE A")
                                       End If
                                    Else
                                       tmpQx1 = 0.0
                                       tmpQx2 = 0.0
                                       VecMaxThis1 = kDur
                                       LivesInd(kDur + 1) = LivesInd(kDur) * tmpLxZ
                                       Exit For
                                    End If
                                 End If
                                 'Under following section, previously interpolated qx's and then applied mortality improvements
                                 'Code changed to apply mortality improvements first and then interpolate qx's between ages
                                 If vIntraYearIsExponential Then

                                    'tmpQx = AgeInterp1 * tmpQx1 + AgeInterp2 * tmpQx2
                                    tmpQx = tmpQx1 'Added Dec 1 2016 (commented out previous line)
                                    tmpWx = AgeInterp1 * tmpWx1 + AgeInterp2 * tmpWx2
                                    tmpLxZ = (1.0 - Math.Max(Math.Min(MortAdd + tmpQx * MortPct + tmpWx, 1), 0)) ^ CalcExponent

                                    If Double.IsNaN(tmpLxZ) OrElse tmpLxZ < 0.0 Then tmpLxZ = 0.0
                                 Else
                                    'Why is there no interpolation here??
                                    'I believe there is no interpolation because the code above looks up a new mortality rate whenever the individual turns an integer age
                                    'As a result, there is no interpolation to be done when Diva looks up a new mortality rate
                                    'But is this right? Why not look up a new rate every twelve months from start of model, ignoring when they turn an integer age?
                                    tmpQx = tmpQx1
                                    tmpWx = tmpWx1

                                    'Line below is commented b/c no interpolation necessary in this case (see above comment)
                                    'tmpQx = AgeInterp1 * tmpQx1 + AgeInterp2 * tmpQx2
                                    tmpDeltaLx = LivesInd(kDur) * Math.Min(1.0, MortAdd + tmpQx * MortPct + tmpWx) * CalcExponent
                                 End If
                              End If
                              If vIntraYearIsExponential Then
                                 LivesInd(kDur + 1) = LivesInd(kDur) * tmpLxZ
                              Else
                                 LivesInd(kDur + 1) = Math.Max(0.0, LivesInd(kDur) - tmpDeltaLx)
                              End If

                           Next kDur

                     End Select
                     If .JointIndex >= 0 Then

                        Select Case ClassList(.classNum - 1, basis).MortPctType
                           Case kMortPctTypeNone
                              MortPct = 1.0
                           Case kMortPctTypeInput
                              If .IsMale Then
                                 MortPct = ClassList(.classNum - 1, basis).MortPctF 'If member is male,  then assume spouse is female
                              Else
                                 MortPct = ClassList(.classNum - 1, basis).MortPctM 'If member is female, then assume spouse is male
                              End If
                           Case kMortPctTypeSeriatim
                              MortPct = .MortPct1
                           Case kMortPctTypeSeriatimM1M2
                              MortPct = .MortPct1
                              MortAdd = .MortPct2
                        End Select

                        tblAgeFractional = .Age2

                        Select Case vAgeBasis
                           Case AgeBasisLast
                              TblAgeX = CInt(Math.Floor(tblAgeFractional))
                           Case AgeBasisNearest
                              TblAgeX = CInt(Math.Round(tblAgeFractional))
                           Case AgeBasisInterpolate
                              TblAge1 = CInt(Math.Floor(tblAgeFractional))
                              TblAge2 = 1 + TblAge1
                              AgeInterp2 = tblAgeFractional - TblAge1
                              AgeInterp1 = 1.0 - AgeInterp2
                           Case AgeBasisInterpolateMonths
                              TblAge1 = CInt(Math.Floor(tblAgeFractional))
                              TblAge2 = 1 + TblAge1
                              AgeInterp2 = Math.Floor(12.0 * (tblAgeFractional - TblAge1)) / 12.0
                              AgeInterp1 = 1.0 - AgeInterp2
                        End Select

                        Select Case vAgeBasis
                           Case AgeBasisLast, AgeBasisNearest
                              PrevDurYr = -1
                              For kDur = MortStartDur To VecMax - 1
                                 'Get the date at this duration, along with the day of year, fraction through year, and calendar year
                                 'These are used for mortality improvement.
                                 currDate = ValDate.AddMonths(kDur * (12 \ CalcFreq))
                                 currDay = currDate.DayOfYear
                                 fracYear = currDay / 365
                                 calYear = currDate.Year
                                 'calYear = ValDate.AddMonths((kDur + 1) * (12 \ CalcFreq)).Year
                                 DurYr = CInt(kDur \ CalcFreq)
                                 If DurYr <> PrevDurYr Then
                                    PrevDurYr = DurYr

                                    If (.IsMale2 And ClassList(.classNum - 1, basis).UseWithdrawalM) Or ((Not .IsMale2) And ClassList(.classNum - 1, basis).UseWithdrawalF) Then
                                       tmpWx = .WxTable2.Wx(TblAgeX)
                                    Else
                                       tmpWx = 0
                                    End If

                                    If DurYr <= .MortTable2.SelectPer Then

                                       If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale2 Then
                                          tmpQxM = tmpQxTableM.Qx(TblAgeX, DurYr)
                                          If ClassList(.classNum - 1, basis).UseImprovementM Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp = (tmpIxTableM.Ix(TblAgeX, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(TblAgeX, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                             tmpQxM *= tmpIxTableM.Ix(TblAgeX, calYear - tmpIxTableM.BaseYear) * incrImp
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale2 Then
                                          tmpQxF = tmpQxTableF.Qx(TblAgeX, DurYr)
                                          If ClassList(.classNum - 1, basis).UseImprovementF Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp = (tmpIxTableF.Ix(TblAgeX, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(TblAgeX, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                             tmpQxF *= tmpIxTableF.Ix(TblAgeX, calYear - tmpIxTableF.BaseYear) * incrImp
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Then
                                          tmpQx = tmpQxM * UniFemPct + tmpQxF * UniMalePct
                                       ElseIf .IsMale2 Then
                                          tmpQx = tmpQxM
                                       ElseIf Not .IsMale2 Then
                                          tmpQx = tmpQxF
                                       Else
                                          FireMsgBox("ERROR CODE A")
                                       End If
                                    Else
                                       tmpAge = TblAgeX + DurYr - .MortTable2.SelectPer
                                       If tmpAge <= .MortTable2.EndAge Then
                                          If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale2 Then
                                             tmpQxM = tmpQxTableM.Qx(tmpAge, tmpQxTableM.SelectPer)
                                             If ClassList(.classNum - 1, basis).UseImprovementM Then
                                                'We need to trend to the exact point in the year
                                                'incrImp is the partial year of improvement needed for the current year
                                                'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                                'Then raise this to the fraction through the year at the current date
                                                incrImp = (tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                                tmpQxM *= tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear) * incrImp
                                             End If
                                          End If

                                          If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale2 Then
                                             tmpQxF = tmpQxTableF.Qx(tmpAge, tmpQxTableF.SelectPer)
                                             If ClassList(.classNum - 1, basis).UseImprovementF Then
                                                'We need to trend to the exact point in the year
                                                'incrImp is the partial year of improvement needed for the current year
                                                'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                                'Then raise this to the fraction through the year at the current date
                                                incrImp = (tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                                tmpQxF *= tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear) * incrImp
                                             End If
                                          End If

                                          If ClassList(.classNum - 1, basis).UseUnisex Then
                                             tmpQx = tmpQxM * UniFemPct + tmpQxF * UniMalePct
                                          ElseIf .IsMale2 Then
                                             tmpQx = tmpQxM
                                          ElseIf Not .IsMale2 Then
                                             tmpQx = tmpQxF
                                          Else
                                             FireMsgBox("ERROR CODE A")
                                          End If

                                       Else
                                          tmpQx = 0.0
                                          VecMaxThis2 = kDur
                                          LivesInd2(kDur + 1) = 0.0
                                          Exit For
                                       End If
                                    End If
                                    If vIntraYearIsExponential Then
                                       tmpLxZ = (1.0 - Math.Max(Math.Min(MortAdd + tmpQx * MortPct + tmpWx, 1), 0)) ^ CalcExponent
                                       If Double.IsNaN(tmpLxZ) OrElse tmpLxZ < 0.0 Then tmpLxZ = 0.0
                                    Else
                                       tmpDeltaLx = LivesInd2(kDur) * (MortAdd + tmpQx * MortPct + tmpWx) * CalcExponent
                                    End If
                                 End If
                                 If vIntraYearIsExponential Then
                                    LivesInd2(kDur + 1) = LivesInd2(kDur) * tmpLxZ
                                 Else
                                    LivesInd2(kDur + 1) = Math.Max(0.0, LivesInd2(kDur) - tmpDeltaLx)
                                 End If
                              Next kDur

                           Case AgeBasisInterpolate, AgeBasisInterpolateMonths
                              PrevDurYr = -1
                              For kDur = MortStartDur To VecMax - 1
                                 'Get the date at this duration, along with the day of year, fraction through year, and calendar year
                                 'These are used for mortality improvement.
                                 currDate = ValDate.AddMonths(kDur * (12 \ CalcFreq))
                                 currDay = currDate.DayOfYear
                                 fracYear = currDay / 365
                                 calYear = currDate.Year
                                 'calYear = ValDate.AddMonths((kDur + 1) * (12 \ CalcFreq)).Year

                                 If vAgeBasis = AgeBasisInterpolateMonths Then
                                    DurYr = CInt(Int((kDur + AgeInterp2 * 12) / CalcFreq))
                                 Else
                                    DurYr = CInt(Int(kDur / CalcFreq))
                                 End If
                                 If DurYr <> PrevDurYr Then
                                    PrevDurYr = DurYr
                                    If (.IsMale2 And ClassList(.classNum - 1, basis).UseWithdrawalM) Or ((Not .IsMale2) And ClassList(.classNum - 1, basis).UseWithdrawalF) Then
                                       tmpWx1 = .WxTable2.Wx(TblAge1)
                                       tmpWx2 = .WxTable2.Wx(TblAge2)
                                    Else
                                       tmpWx1 = 0
                                       tmpWx2 = 0
                                    End If

                                    If DurYr <= .MortTable2.SelectPer Then
                                       If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale2 Then
                                          tmpQx1M = tmpQxTableM.Qx(TblAge1, DurYr)
                                          tmpQx2M = tmpQxTableM.Qx(TblAge2, DurYr)
                                          If ClassList(.classNum - 1, basis).UseImprovementM Then
                                             'We need to trend to the exact point in the year
                                             'incrImp is the partial year of improvement needed for the current year
                                             'We divide the compounded improvement factors to get the (uncompounded) improvement rate for the current year
                                             'Then raise this to the fraction through the year at the current date
                                             incrImp1 = (tmpIxTableM.Ix(TblAge1, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(TblAge1, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                             incrImp2 = (tmpIxTableM.Ix(TblAge2, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(TblAge2, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                             tmpQx1M *= tmpIxTableM.Ix(TblAge1, calYear - tmpIxTableM.BaseYear) * incrImp1
                                             tmpQx2M *= tmpIxTableM.Ix(TblAge2, calYear - tmpIxTableM.BaseYear) * incrImp2
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale2 Then
                                          tmpQx1F = tmpQxTableF.Qx(TblAge1, DurYr)
                                          tmpQx2F = tmpQxTableF.Qx(TblAge2, DurYr)
                                          If ClassList(.classNum - 1, basis).UseImprovementF Then
                                             incrImp1 = (tmpIxTableF.Ix(TblAge1, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(TblAge1, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                             incrImp2 = (tmpIxTableF.Ix(TblAge2, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(TblAge2, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                             tmpQx1F *= tmpIxTableF.Ix(TblAge1, calYear - tmpIxTableF.BaseYear) * incrImp1
                                             tmpQx2F *= tmpIxTableF.Ix(TblAge2, calYear - tmpIxTableF.BaseYear) * incrImp2
                                          End If
                                       End If

                                       If ClassList(.classNum - 1, basis).UseUnisex Then
                                          tmpQx1 = tmpQx1M * UniFemPct + tmpQx1F * UniMalePct
                                          tmpQx2 = tmpQx2M * UniFemPct + tmpQx2F * UniMalePct
                                       ElseIf .IsMale2 Then
                                          tmpQx1 = tmpQx1M
                                          tmpQx2 = tmpQx2M
                                       ElseIf Not .IsMale2 Then
                                          tmpQx1 = tmpQx1F
                                          tmpQx2 = tmpQx2F
                                       Else
                                          FireMsgBox("ERROR CODE A")
                                       End If
                                    Else
                                       tmpAge = TblAge1 + DurYr - .MortTable2.SelectPer
                                       If tmpAge < .MortTable2.EndAge Then '   Note use of < rather than <= due to interpolation
                                          If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale2 Then
                                             tmpQx1M = tmpQxTableM.Qx(tmpAge, tmpQxTableM.SelectPer)
                                             tmpQx2M = tmpQxTableM.Qx(tmpAge + 1, tmpQxTableM.SelectPer)
                                             If ClassList(.classNum - 1, basis).UseImprovementM Then
                                                incrImp1 = (tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                                incrImp2 = (tmpIxTableM.Ix(tmpAge + 1, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(tmpAge + 1, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                                tmpQx1M *= tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear) * incrImp1
                                                tmpQx2M *= tmpIxTableM.Ix(tmpAge + 1, calYear - tmpIxTableM.BaseYear) * incrImp2
                                             End If
                                          End If

                                          If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale2 Then
                                             tmpQx1F = tmpQxTableF.Qx(tmpAge, tmpQxTableF.SelectPer)
                                             tmpQx2F = tmpQxTableF.Qx(tmpAge + 1, tmpQxTableF.SelectPer)
                                             If ClassList(.classNum - 1, basis).UseImprovementF Then
                                                incrImp1 = (tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                                incrImp2 = (tmpIxTableF.Ix(tmpAge + 1, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(tmpAge + 1, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                                tmpQx1F *= tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear) * incrImp1
                                                tmpQx2F *= tmpIxTableF.Ix(tmpAge + 1, calYear - tmpIxTableF.BaseYear) * incrImp2
                                             End If
                                          End If

                                          If ClassList(.classNum - 1, basis).UseUnisex Then
                                             tmpQx1 = tmpQx1M * UniFemPct + tmpQx1F * UniMalePct
                                             tmpQx2 = tmpQx2M * UniFemPct + tmpQx2F * UniMalePct
                                          ElseIf .IsMale2 Then
                                             tmpQx1 = tmpQx1M
                                             tmpQx2 = tmpQx2M
                                          ElseIf Not .IsMale2 Then
                                             tmpQx1 = tmpQx1F
                                             tmpQx2 = tmpQx2F
                                          Else
                                             FireMsgBox("ERROR CODE A")
                                          End If
                                       ElseIf tmpAge = .MortTable2.EndAge Then     '   Assume Final Qx is 1.00
                                          If ClassList(.classNum - 1, basis).UseUnisex Or .IsMale2 Then
                                             tmpQx1M = tmpQxTableM.Qx(tmpAge, tmpQxTableM.SelectPer)
                                             tmpQx2M = 1
                                             If ClassList(.classNum - 1, basis).UseImprovementM Then
                                                incrImp1 = (tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear + 1) / tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear)) ^ fracYear
                                                tmpQx1M *= tmpIxTableM.Ix(tmpAge, calYear - tmpIxTableM.BaseYear) * incrImp1
                                                tmpQx2M *= 1
                                             End If
                                          End If

                                          If ClassList(.classNum - 1, basis).UseUnisex Or Not .IsMale2 Then
                                             tmpQx1F = tmpQxTableF.Qx(tmpAge, tmpQxTableF.SelectPer)
                                             tmpQx2F = 1
                                             If ClassList(.classNum - 1, basis).UseImprovementF Then
                                                incrImp1 = (tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear + 1) / tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear)) ^ fracYear
                                                tmpQx1F *= tmpIxTableF.Ix(tmpAge, calYear - tmpIxTableF.BaseYear) * incrImp1
                                                tmpQx2F *= 1
                                             End If
                                          End If

                                          If ClassList(.classNum - 1, basis).UseUnisex Then
                                             tmpQx1 = tmpQx1M * UniFemPct + tmpQx1F * UniMalePct
                                             tmpQx2 = tmpQx2M * UniFemPct + tmpQx2F * UniMalePct
                                          ElseIf .IsMale2 Then
                                             tmpQx1 = tmpQx1M
                                             tmpQx2 = tmpQx2M
                                          ElseIf Not .IsMale2 Then
                                             tmpQx1 = tmpQx1F
                                             tmpQx2 = tmpQx2F
                                          Else
                                             FireMsgBox("ERROR CODE A")
                                          End If
                                       Else
                                          tmpQx1 = 0.0
                                          tmpQx2 = 0.0
                                          VecMaxThis2 = kDur
                                          LivesInd2(kDur + 1) = LivesInd2(kDur) * tmpLxZ
                                          Exit For
                                       End If
                                    End If

                                    If vIntraYearIsExponential Then
                                       'tmpQx = AgeInterp1 * tmpQx1 + AgeInterp2 * tmpQx2
                                       tmpQx = tmpQx1 'Added Dec 1 2016 (commented out previous line)
                                       tmpWx = AgeInterp1 * tmpWx1 + AgeInterp2 * tmpWx2
                                       tmpLxZ = (1.0 - Math.Max(Math.Min(MortAdd + tmpQx * MortPct + tmpWx, 1), 0)) ^ CalcExponent

                                       If Double.IsNaN(tmpLxZ) OrElse tmpLxZ < 0.0 Then tmpLxZ = 0.0
                                    Else
                                       'tmpQx = AgeInterp1 * tmpQx1 + AgeInterp2 * tmpQx2
                                       tmpQx = tmpQx1 'Added Dec 1 2016 (commented out previous line)
                                       tmpWx = AgeInterp1 * tmpWx1 + AgeInterp2 * tmpWx2
                                       tmpDeltaLx = LivesInd2(kDur) * Math.Min(1.0, MortAdd + tmpQx * MortPct + tmpWx) * CalcExponent
                                    End If
                                 End If
                                 If vIntraYearIsExponential Then
                                    LivesInd2(kDur + 1) = LivesInd2(kDur) * tmpLxZ
                                 Else
                                    LivesInd2(kDur + 1) = Math.Max(0.0, LivesInd2(kDur) - tmpDeltaLx)
                                 End If
                              Next kDur
                        End Select

                        Select Case .JointIndex
                           Case kJointReduceEither
                              '   Benefits reduce on EITHER death
                              tmpFctr = 2.0 * .SurvivorPct - 1.0
                              For kDur = MortStartDur To Math.Min(VecMaxThis1, VecMaxThis2)
                                 LivesInd(kDur) = .SurvivorPct * (LivesInd(kDur) + LivesInd2(kDur)) - tmpFctr * LivesInd2(kDur) * LivesInd(kDur)
                              Next kDur
                              If VecMaxThis1 > VecMaxThis2 Then
                                 For kDur = VecMaxThis2 + 1 To VecMaxThis1
                                    LivesInd(kDur) = .SurvivorPct * LivesInd(kDur)
                                 Next kDur
                              Else
                                 For kDur = VecMaxThis1 + 1 To VecMaxThis2
                                    LivesInd(kDur) = .SurvivorPct * LivesInd2(kDur)
                                 Next kDur
                                 VecMaxThis1 = VecMaxThis2
                              End If
                           Case kJointReduceLife1
                              '   Benefits reduce on death of Life 1
                              For kDur = MortStartDur To Math.Min(VecMaxThis1, VecMaxThis2)
                                 LivesInd(kDur) = LivesInd(kDur) + .SurvivorPct * LivesInd2(kDur) - .SurvivorPct * LivesInd2(kDur) * LivesInd(kDur)
                              Next kDur
                              If VecMaxThis1 > VecMaxThis2 Then
                                 For kDur = VecMaxThis2 + 1 To VecMaxThis1
                                    LivesInd(kDur) = LivesInd(kDur)
                                 Next kDur
                              Else
                                 For kDur = VecMaxThis1 + 1 To VecMaxThis2
                                    LivesInd(kDur) = .SurvivorPct * LivesInd2(kDur)
                                 Next kDur
                                 VecMaxThis1 = VecMaxThis2
                              End If
                           Case kJointReduceLife2
                              '   Benefits reduce on death of Life 2
                              For kDur = MortStartDur To Math.Min(VecMaxThis1, VecMaxThis2)
                                 LivesInd(kDur) = .SurvivorPct * LivesInd(kDur) + LivesInd2(kDur) - .SurvivorPct * LivesInd2(kDur) * LivesInd(kDur)
                              Next kDur
                              If VecMaxThis1 > VecMaxThis2 Then
                                 For kDur = VecMaxThis2 + 1 To VecMaxThis1
                                    LivesInd(kDur) = .SurvivorPct * LivesInd(kDur)
                                 Next kDur
                              Else
                                 For kDur = VecMaxThis1 + 1 To VecMaxThis2
                                    LivesInd(kDur) = LivesInd2(kDur)
                                 Next kDur
                                 VecMaxThis1 = VecMaxThis2
                              End If
                        End Select
                     End If

                     AnnSumPost = 0.0
                     prevLivesInd = 0.0
                     Dim tmpT As Integer
                     tmpT = Math.Min(VecMaxThis1 + 1, npvAnnIndPost.GetUpperBound(0)) 'npvPremIndPost.GetUpperBound(0)) *****
                     npvAnnIndPost(tmpT) = 0.0
                     npvAnnIndNormalAge(tmpT) = 0.0

                     'Starting Annuity Factor calculations + Retirement Age Loop
                     Dim kStart, kEnd, tmpRetireDur, tmpGteeDur As Integer
                     Dim PensionMaxValue(VecMax), PensionMaxCF(VecMax), AgeOfMaxValue As Double
                     Dim tmpPvPension, tmpAnnBenefit As Double
                     Dim normalAgeInd As Boolean = False
                     Dim tmpLxRetire As Double = 1
                     tmpAnnBenefit = .AnnBenefit

                     If UseRetireLoop Then
                        kStart = 0
                        kEnd = ClassList(.classNum - 1, basis).RetLoopAgeFactors.GetUpperBound(0)
                     Else
                        kStart = 0
                        kEnd = 0
                     End If
                     PensionMaxValue(0) = Double.MinValue
                     AgeOfMaxValue = 0.0

                     For k = 0 To VecMax
                        npvAnnSumNormalAge(k) = 0.0
                     Next

                     For RetireLoop As Integer = kStart To kEnd

                        'Zero out sum variables
                        For k = 0 To VecMax
                           AnnSumPost = 0.0
                           AnnPmtSumPost(k) = 0.0
                           npvAnnSumPost(k) = 0.0
                        Next

                        If UseRetireLoop Then
                           tmpRetireDur = CInt(CalcFreq * (ClassList(.classNum - 1, basis).RetLoopAgeFactors(RetireLoop, 0) - .Age1))
                           If tmpRetireDur < 0 Then tmpRetireDur = 0 '11/27/2016 -- no 'tmp' before. Added it in. I think this is why immediate retirement was not working before.

                           'Added 08/17/2016
                           If UseMortPreRetire Then
                              tmpLxRetire = 1
                           Else
                              tmpLxRetire = LivesInd(Math.Max(tmpRetireDur, 0))
                           End If
                           ''

                           tmpAnnBenefit = .AnnBenefit * ClassList(.classNum - 1, basis).RetLoopAgeFactors(RetireLoop, 1)
                           If ClassList(.classNum - 1, basis).RetLoopAgeFactors(RetireLoop, 0) = NormalRetireAge Then
                              normalAgeInd = True
                           Else
                              normalAgeInd = False
                           End If
                        Else
                           tmpRetireDur = CInt(CalcFreq * (.RetireAgeSeriatim - .Age1))
                           If tmpRetireDur < 0 Then tmpRetireDur = 0 'Added 11/27/2016
                           normalAgeInd = True
                        End If

                        tmpGteeDur = tmpRetireDur + CInt(.LastGteeYears * CalcFreq)
                        tmpPvPension = 0.0
                        For kDur = VecMaxThis1 To 0 Step -1
                           If tmpRetireDur <= kDur Then
                              '                              If kDur < GteeAnnDur Then
                              If kDur < tmpGteeDur Then
                                 AnnPmtInd(kDur) = LivesInd(tmpRetireDur) * tmpAnnBenefit * ClassList(.classNum - 1, basis).InflFctr(kDur) / tmpLxRetire
                              Else
                                 Select Case vCashFlowAnnuity
                                    Case vCashFlowStart
                                       AnnPmtInd(kDur) = LivesInd(kDur) * tmpAnnBenefit * ClassList(.classNum - 1, basis).InflFctr(kDur) / tmpLxRetire
                                    Case vCashFlowMiddle
                                       AnnPmtInd(kDur) = 0.5 * (prevLivesInd + LivesInd(kDur)) * tmpAnnBenefit * ClassList(.classNum - 1, basis).InflFctr(kDur) / tmpLxRetire
                                    Case vCashFlowEnd
                                       AnnPmtInd(kDur) = prevLivesInd * tmpAnnBenefit * ClassList(.classNum - 1, basis).InflFctr(kDur) / tmpLxRetire
                                 End Select
                              End If
                           Else
                              AnnPmtInd(kDur) = 0.0
                           End If
                           prevLivesInd = LivesInd(kDur)

                           If kDur < tmpRetireDur Then

                           Else
                              AnnPmtSumPost(kDur) += AnnPmtInd(kDur)
                              AnnSumPost += AnnPmtInd(kDur) * ClassList(.classNum - 1, basis).IntrFctrAnn(kDur)
                           End If

                           npvAnnIndPost(kDur) = AnnSumPost / ClassList(.classNum - 1, basis).IntrFctrStart(kDur)
                           npvAnnSumPost(kDur) += npvAnnIndPost(kDur)

                           'LivesSum(kDur) += LivesInd(kDur)
                           If kDur = tmpRetireDur Then
                              RetireAnn = AnnSumPost
                           End If

                           If normalAgeInd = True Then
                              npvAnnIndNormalAge(kDur) = npvAnnIndPost(kDur)
                              npvAnnSumNormalAge(kDur) = npvAnnSumPost(kDur)

                              npvAnnClass(.classNum - 1, basis, kDur) += npvAnnIndPost(kDur)
                              cfAnnClass(.classNum - 1, basis, kDur) += AnnPmtInd(kDur)

                              LivesSum(basis, kDur) += LivesInd(kDur)
                           End If

                        Next kDur

                        If UseRetireLoop Then
                           tmpPvPension = npvAnnSumPost(0)

                           If tmpPvPension > PensionMaxValue(0) Then
                              For kDur = VecMaxThis1 To 0 Step -1
                                 PensionMaxValue(kDur) = npvAnnSumPost(kDur)  'Would this still work if it was changed to npvAnnIndPost??
                                 PensionMaxCF(kDur) = AnnPmtInd(kDur)
                              Next
                              AgeOfMaxValue = ClassList(.classNum - 1, basis).RetLoopAgeFactors(RetireLoop, 0)
                           End If

                        End If

                     Next RetireLoop

                     If UseRetireLoop Then
                        For kDur = VecMaxThis1 To 0 Step -1
                           npvBestAgeClass(.classNum - 1, basis, kDur) += PensionMaxValue(kDur)
                           cfBestAgeClass(.classNum - 1, basis, kDur) += PensionMaxCF(kDur)
                        Next
                     End If

                     If HasAnySeriatimOut Or HasAnySupOut Then
                        For tpSeria = 0 To Math.Max(0, TimePeriodsToDump)
                           tpToUse = tpSeria * CalcFreq \ CInt(gv.TimePeriodsPerYear)
                           For k = 0 To SeriatimCol.GetUpperBound(0)
                              If SeriatimCol(k) <= sMaxInCol Then
                                 SeriatimOutLine(k) = .Input(SeriatimCol(k))
                                 Select Case SeriatimCol(k)
                                    Case sColInterest
                                       Select Case ClassList(.classNum - 1, basis).IntrMethod
                                          Case kMethodRate
                                             SeriatimOutLine(k) = ParameterInterestExpr
                                          Case kMethodVector
                                             SeriatimOutLine(k) = ParameterInterestVector
                                          Case kMethodSeriatim, kMethodSeriatimVector
                                             '   No change
                                       End Select

                                    Case sColMaxGteeDate
                                       Select Case LastGteeType
                                          Case kLastGteeNone
                                             SeriatimOutLine(k) = "None"
                                          Case kLastGteeDate, kLastGteeYears
                                             '   No change
                                          Case kLastGteeFixed
                                             SeriatimOutLine(k) = ParameterGteeYrs
                                       End Select
                                 End Select
                              Else
                                 Select Case basis
                                    Case 0
                                       Select Case SeriatimCol(k)
                                          Case kOutAnnGC
                                             tmpOutVal = npvAnnSumNormalAge(tpToUse)
                                             SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                             If tpSeria = 0 Then SumOuts(k) += tmpOutVal
                                          Case kOutBestAgeGC
                                             If UseRetireLoop Then
                                                tmpOutVal = AgeOfMaxValue
                                                SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                             Else
                                                tmpOutVal = Nothing
                                                SeriatimOutLine(k) = "N/A"
                                             End If
                                          Case kOutBestAnnGC
                                             If UseRetireLoop Then
                                                tmpOutVal = PensionMaxValue(tpToUse)
                                                SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                                If tpSeria = 0 Then SumOuts(k) += tmpOutVal
                                             Else
                                                tmpOutVal = Nothing
                                                SeriatimOutLine(k) = "N/A"
                                             End If

                                       End Select
                                    Case 1
                                       Select Case SeriatimCol(k)
                                          Case kOutAnnSY
                                             tmpOutVal = npvAnnSumNormalAge(tpToUse)
                                             SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                             If tpSeria = 0 Then SumOuts(k) += tmpOutVal
                                          Case kOutBestAgeSY
                                             If UseRetireLoop Then
                                                tmpOutVal = AgeOfMaxValue
                                                SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                             Else
                                                tmpOutVal = Nothing
                                                SeriatimOutLine(k) = "N/A"
                                             End If

                                          Case kOutBestAnnSY
                                             If UseRetireLoop Then
                                                tmpOutVal = PensionMaxValue(tpToUse)
                                                SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                                If tpSeria = 0 Then SumOuts(k) += tmpOutVal
                                             Else
                                                tmpOutVal = Nothing
                                                SeriatimOutLine(k) = "N/A"
                                             End If
                                       End Select
                                    Case 2
                                       Select Case SeriatimCol(k)
                                          Case kOutAnnOT
                                             tmpOutVal = npvAnnSumNormalAge(tpToUse)
                                             SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                             If tpSeria = 0 Then SumOuts(k) += tmpOutVal
                                          Case kOutBestAgeOT
                                             If UseRetireLoop Then
                                                tmpOutVal = AgeOfMaxValue
                                                SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                             Else
                                                tmpOutVal = Nothing
                                                SeriatimOutLine(k) = "N/A"
                                             End If
                                          Case kOutBestAnnOT
                                             If UseRetireLoop Then
                                                tmpOutVal = PensionMaxValue(tpToUse)
                                                SeriatimOutLine(k) = Format(tmpOutVal, vNumFmt)
                                                If tpSeria = 0 Then SumOuts(k) += tmpOutVal
                                             Else
                                                tmpOutVal = Nothing
                                                SeriatimOutLine(k) = "N/A"
                                             End If

                                       End Select
                                 End Select
                              End If
                           Next k
                           If HasAnySeriatimOut AndAlso tpSeria = 0 Then
                              SeriatimOut(kItem) = Join(SeriatimOutLine, vbTab)
                           End If
                           If HasAnySupOut AndAlso tpSeria <= TimePeriodsToDump Then
                              SuppOut(tpSeria)(kItem) = Join(SeriatimOutLine, vbTab)
                           End If
                        Next tpSeria
                     Else
                        '   Nothing
                     End If

                     '*****************************************

                     '   Time Zero Present Values

                     SumAnnPv = npvAnnIndPost(0)

                     '*****************************************

                     '   Do sums across time periods for age-buckets
                     tpS = 0
                     Dim tpMax As Integer
                     If NeedTimeSum Then
                        For tp = 1 To gv.TimeSeriesMax
                           tpE = Math.Min(CInt(NumTimeItems * tp), VecMaxThis1)
                           If tpE <= tpS Then Exit For
                           SumAnn = 0.0
                           For tX = tpS To tpE - 1
                              SumAnn += AnnPmtInd(tX)
                           Next tX
                           SumAnnPv = npvAnnIndPost(tpE)

                           tpS = tpE
                        Next tp
                     Else
                        tpMax = Math.Min(gv.TimeSeriesMax, VecMaxThis1 + 1)

                        For tp = 1 To tpMax
                           tpE = tp - 1
                           SumAnnPv = npvAnnIndPost(tp)
                        Next tp
                     End If
                  End With
               Catch ex As Exception
                  WarnCount += 1
                  If WarnCount <= 10 Then
                     FireMsgBox("In node '" & vNodeName & "', error here. Item: " & CStr(kItem) & ", Duration " & CStr(kDur) & vbCrLf & "   " & ex.Message)
                     If TypeOf ex Is System.IndexOutOfRangeException Then
                        FireMsgBox("   Probably you are using ages outside the range of your table.  Extend your mortality tables to cover all ages of seriatim data.")
                     End If
                  End If
                  If WarnCount = 10 Then
                     FireMsgBox(vbCrLf & "In Node '" & vNodeName & "', 10 warnings reached.  Further messages suppressed.")
                  End If

               End Try
            Next basis
         Next kItem

         If HasAnySeriatimOut Then
            Dim tmpStr As String = SeriatimOutTotalSpacer
            For k = 0 To SumOuts.GetUpperBound(0)
               If SumOuts(k) = 0.0 Then
                  tmpStr &= vbTab
               Else
                  tmpStr &= Format(SumOuts(k), vNumFmt) & vbTab
               End If
            Next k
            SeriatimOut(SeriaMax + 1) = tmpStr.Substring(0, tmpStr.Length - 1)
         End If

         '   Dump aggregate payouts into elements
         tpS = 0


         'vEleVal(sElePvAnnuityEnd, 0) = npvAnnSumPost(0)

         Dim numClasses As Integer = ClassList.GetUpperBound(0) + 1

         For tp = 0 To gv.TimeSeriesMax
            For kClass As Integer = 0 To numClasses - 1

               'Class-level present value outputs for normal retirement age
               vEleVal(sMaxElementsBase + 1 + oClassElePV_GC * numClasses + kClass, tp) = npvAnnClass(kClass, 0, tp * intNumTimeItems)
               vEleVal(sMaxElementsBase + 1 + oClassElePV_SY * numClasses + kClass, tp) = npvAnnClass(kClass, 1, tp * intNumTimeItems)
               vEleVal(sMaxElementsBase + 1 + oClassElePV_OT * numClasses + kClass, tp) = npvAnnClass(kClass, 2, tp * intNumTimeItems)
               'Class-level present value outputs for best retirement age
               vEleVal(sMaxElementsBase + 1 + oClassEleBestPV_GC * numClasses + kClass, tp) = npvBestAgeClass(kClass, 0, tp * intNumTimeItems)
               vEleVal(sMaxElementsBase + 1 + oClassEleBestPV_SY * numClasses + kClass, tp) = npvBestAgeClass(kClass, 1, tp * intNumTimeItems)
               vEleVal(sMaxElementsBase + 1 + oClassEleBestPV_OT * numClasses + kClass, tp) = npvBestAgeClass(kClass, 2, tp * intNumTimeItems)
               'Total present value outputs for normal retirement age
               vEleVal(sEleTotalPV_NRA_GC, tp) += npvAnnClass(kClass, 0, tp * intNumTimeItems)
               vEleVal(sEleTotalPV_NRA_SY, tp) += npvAnnClass(kClass, 1, tp * intNumTimeItems)
               vEleVal(sEleTotalPV_NRA_OT, tp) += npvAnnClass(kClass, 2, tp * intNumTimeItems)
               'Total present value outputs for best retirement age
               vEleVal(sEleTotalPV_MaxCV_GC, tp) += npvBestAgeClass(kClass, 0, tp * intNumTimeItems)
               vEleVal(sEleTotalPV_MaxCV_SY, tp) += npvBestAgeClass(kClass, 1, tp * intNumTimeItems)
               vEleVal(sEleTotalPV_MaxCV_OT, tp) += npvBestAgeClass(kClass, 2, tp * intNumTimeItems)
            Next kClass
         Next tp


         'Now we do cash flow outputs. This needs to be done separately because if model frequency is smaller than calculation frequency (e.g. model is annual but calculations are monthly),
         'then we need to sum up the cash flows to line up the time periods. For example, if model is annual and calculations are monthly, we need to add up the cash flows for Jan-Dec to get
         'the annual cash flow. The same is not true of present values because present values are "snapshots in time."
         'We also do outputs for number of lives/deaths.

         'Set the ending number of lives for time period 0:
         vEleVal(sEleLivesEnd_GC, 0) = LivesSum(0, 0)
         vEleVal(sEleLivesEnd_SY, 0) = LivesSum(1, 0)
         vEleVal(sEleLivesEnd_OT, 0) = LivesSum(2, 0)

         'Now the real stuff:
         For kClass As Integer = 0 To numClasses - 1
            For kBasis As Integer = 0 To 2
               tpS = 0  'tpStart -- time period to start sum (e.g. beginning of year)
               For tp = 1 To gv.TimeSeriesMax

                  If NeedTimeSum Then
                     SumCF = 0.0
                     SumBestCF = 0.0
                     SumAnnPost = 0.0
                     tpE = CInt(NumTimeItems * tp) 'tpEnd -- time period to end sum (e.g. end of year)
                     For t As Integer = tpS To tpE - 1
                        SumCF += cfAnnClass(kClass, kBasis, t)
                        SumBestCF += cfBestAgeClass(kClass, kBasis, t)
                     Next

                     Select Case kBasis
                        Case 0
                           vEleVal(sMaxElementsBase + 1 + oClassEleCF_GC * numClasses + kClass, tp) = SumCF
                           vEleVal(sMaxElementsBase + 1 + oClassEleBestCF_GC * numClasses + kClass, tp) = SumBestCF
                           vEleVal(sEleTotalCF_NRA_GC, tp) += SumCF
                           vEleVal(sEleTotalCF_MaxCV_GC, tp) += SumBestCF
                        Case 1
                           vEleVal(sMaxElementsBase + 1 + oClassEleCF_SY * numClasses + kClass, tp) = SumCF
                           vEleVal(sMaxElementsBase + 1 + oClassEleBestCF_SY * numClasses + kClass, tp) = SumBestCF
                           vEleVal(sEleTotalCF_NRA_SY, tp) += SumCF
                           vEleVal(sEleTotalCF_MaxCV_SY, tp) += SumBestCF
                        Case 2
                           vEleVal(sMaxElementsBase + 1 + oClassEleCF_OT * numClasses + kClass, tp) = SumCF
                           vEleVal(sMaxElementsBase + 1 + oClassEleBestCF_OT * numClasses + kClass, tp) = SumBestCF
                           vEleVal(sEleTotalCF_NRA_OT, tp) += SumCF
                           vEleVal(sEleTotalCF_MaxCV_OT, tp) += SumBestCF
                     End Select
                  Else
                     'Maybe we don't need to sum across time periods at all (i.e. maybe model frequency = calculation frequency). This is handled below (it's simpler).
                     tpS = tp - 1
                     tpE = tp

                     Select Case kBasis
                        Case 0
                           vEleVal(sMaxElementsBase + 1 + oClassEleCF_GC * numClasses + kClass, tp) = cfAnnClass(kClass, kBasis, tpS)
                           vEleVal(sMaxElementsBase + 1 + oClassEleBestCF_GC * numClasses + kClass, tp) = cfBestAgeClass(kClass, kBasis, tpS)
                           vEleVal(sEleTotalCF_NRA_GC, tp) += cfAnnClass(kClass, kBasis, tpS)
                           vEleVal(sEleTotalCF_MaxCV_GC, tp) += cfBestAgeClass(kClass, kBasis, tpS)
                        Case 1
                           vEleVal(sMaxElementsBase + 1 + oClassEleCF_SY * numClasses + kClass, tp) = cfAnnClass(kClass, kBasis, tpS)
                           vEleVal(sMaxElementsBase + 1 + oClassEleBestCF_SY * numClasses + kClass, tp) = cfBestAgeClass(kClass, kBasis, tpS)
                           vEleVal(sEleTotalCF_NRA_SY, tp) += cfAnnClass(kClass, kBasis, tpS)
                           vEleVal(sEleTotalCF_MaxCV_SY, tp) += cfBestAgeClass(kClass, kBasis, tpS)
                        Case 2
                           vEleVal(sMaxElementsBase + 1 + oClassEleCF_OT * numClasses + kClass, tp) = cfAnnClass(kClass, kBasis, tpS)
                           vEleVal(sMaxElementsBase + 1 + oClassEleBestCF_OT * numClasses + kClass, tp) = cfBestAgeClass(kClass, kBasis, tpS)
                           vEleVal(sEleTotalCF_NRA_OT, tp) += cfAnnClass(kClass, kBasis, tpS)
                           vEleVal(sEleTotalCF_MaxCV_OT, tp) += cfBestAgeClass(kClass, kBasis, tpS)
                     End Select
                  End If

                  vEleVal(sEleLivesStart_GC, tp) = LivesSum(0, tpS)
                  vEleVal(sEleLivesEnd_GC, tp) = LivesSum(0, tpE)
                  vEleVal(sEleDeaths_GC, tp) = LivesSum(0, tpS) - LivesSum(0, tpE)

                  vEleVal(sEleLivesStart_SY, tp) = LivesSum(1, tpS)
                  vEleVal(sEleLivesEnd_SY, tp) = LivesSum(1, tpE)
                  vEleVal(sEleDeaths_SY, tp) = LivesSum(1, tpS) - LivesSum(1, tpE)

                  vEleVal(sEleLivesStart_OT, tp) = LivesSum(2, tpS)
                  vEleVal(sEleLivesEnd_OT, tp) = LivesSum(2, tpE)
                  vEleVal(sEleDeaths_OT, tp) = LivesSum(2, tpS) - LivesSum(2, tpE)

                  tpS = tpE
               Next tp
            Next kBasis
         Next kClass

         'For tp = 1 To gv.TimeSeriesMax
         '	'If NeedTimeSum Then
         '	'    SumAnnPost = 0.0
         '	'    tpE = CInt(NumTimeItems * tp)
         '	'    For tX = tpS To tpE - 1
         '	'        SumAnnPost += AnnPmtSumPost(tX)
         '	'    Next tX
         '	'    vEleVal(sEleAnnuity, tp) = SumAnnPost
         '	'Else
         '	'    tpS = tp - 1
         '	'    tpE = tp
         '	'    vEleVal(sEleAnnuity, tp) = AnnPmtSumPost(tpS)
         '	'End If



         '	'vEleVal(sElePvAnnuityStrt, tp) = npvAnnSumPost(tpS)
         '	'vEleVal(sElePvAnnuityEnd, tp) = npvAnnSumPost(tpE)

         '	tpS = tpE
         'Next tp

         If gv.ScenarioNum = 1 Then '   Upgrade property form once only, not on each thread
            If HasAnySeriatimOut Then
               vParamArray(sTabSeriatimOut) = Join(SeriatimOut, vbCr)
            Else
               vParamArray(sTabSeriatimOut) = vbTab & vbCr & vbTab
            End If
            If ParentNode.HasPropertyForm AndAlso vPropertyTabIndex = sTabSeriatimOut Then
               MyBase.FireRefreshEvent(0)
            End If
         End If

         If HasAnySupOut Then
            If gv.ScenarioNum <= ScenariosToDump Then
               For k = 0 To TimePeriodsToDump
                  vSupIO.WriteLine(vbCrLf & vbCrLf & "SCENARIO " & Format(gv.ScenarioNum, "N0") & ", TIME PERIOD " & Format(k, "N0") & ", " & gv.TimeSeriesFmted(k) & vbCrLf)
                  vSupIO.WriteLine(OutColHeading & vbCrLf)
                  vSupIO.WriteLine(Join(SuppOut(k), vbCrLf))
               Next k
               vSupIO.Flush()
            End If
         End If

      End If

      IsFirstScenario = False
      If HasAnySupOut Then
         If gv.ScenarioNum = ScenariosToDump And gv.TimePeriod = gv.TimeSeriesMax Then
            Try
               vSupIO.Close()
            Catch ex As Exception
               '   Ignore
            End Try
         End If
      End If

   End Sub

#End Region

#Region "Other Utilities"

   Private Sub GetSeriatimData(basis As Integer)
      '''This sub gets all seriatim data and stores it in the "classlist" variable.
      ''' Input is the basis


   End Sub

   Private Function GetSeriatimDate(ByVal InputStr As String) As Date
      '   Interpret seriatim Age or DOB
      InputStr = InputStr.PadRight(DateLen)
      Select Case DateType
         Case 0, 1   '   Treat as VB date expression
            GetSeriatimDate = CDate(InputStr)
         Case 2     '   Has no Y, M or D
            GetSeriatimDate = ValDate
         Case 3     '   Has YYYY but no M or D
            GetSeriatimDate = DateSerial(CInt(InputStr.Substring(yStart, yLen)), 1, 1)
         Case 4     '   Has YY but no M or D
            GetSeriatimDate = DateSerial(1900 + CInt(InputStr.Substring(yStart, yLen)), 1, 1)
         Case 5     '   Has YYYY, M but no D
            GetSeriatimDate = DateSerial(CInt(InputStr.Substring(yStart, yLen)), CInt(InputStr.Substring(mStart, mLen)), 15)
         Case 6     '   Has YY, M but no D
            GetSeriatimDate = DateSerial(1900 + CInt(InputStr.Substring(yStart, yLen)), CInt(InputStr.Substring(mStart, mLen)), 15)
         Case 7     '   Has YYYY, M, D
            GetSeriatimDate = DateSerial(CInt(InputStr.Substring(yStart, yLen)), _
               CInt(InputStr.Substring(mStart, mLen)), CInt(InputStr.Substring(dStart, dLen)))
         Case 8     '   Has YY, M, D
            GetSeriatimDate = DateSerial(1900 + CInt(InputStr.Substring(yStart, yLen)), _
               CInt(InputStr.Substring(mStart, mLen)), CInt(InputStr.Substring(dStart, dLen)))
      End Select

   End Function

   Private Sub InterpretVector(ByRef InString As String, ByRef OutVec() As Double)
      Dim kItem, kYr, kPrev As Integer
      Dim InList As String(), tmpVal As Double
      ReDim OutVec(VecMax)
      If InString.Trim = "" Then
         FireMsgBox("In '" & vNodeName & "', You have specified a vector but haven't provided it.")
         Exit Sub
      End If
      InList = ParseToStr(InString)
      kPrev = -1
      If InList.GetUpperBound(0) >= 1 Then
         For kItem = 0 To InList.GetUpperBound(0) - 1 Step 2
            tmpVal = DivaCalc.EvaluateMathExpression(InList(kItem))
            For kYr = kPrev + 1 To kPrev + CalcFreq * CInt(InList(kItem + 1))
               OutVec(kYr) = tmpVal
            Next kYr
            kPrev = kPrev + CalcFreq * CInt(InList(kItem + 1))
         Next kItem
      End If
      If (InList.GetUpperBound(0) Mod 2) = 0 Then
         tmpVal = DivaCalc.EvaluateMathExpression(InList(InList.GetUpperBound(0)))
      End If
      For kYr = kPrev + 1 To VecMax
         OutVec(kYr) = tmpVal
      Next kYr

   End Sub

   Private Function ReadWxSheets() As Boolean
      '   Read in a table and store results in Public Variable QxTable(Sex,Smk)
      Dim t As Integer
      Dim kRow As Integer, SheetName As String
      Dim IsMatrix As Boolean
      Dim gVectorSet As DivaCalcTools.Structures.VectorSet = Nothing
      Dim tmpArray(,) As Double

      'Get list of withdrawal table names
      Dim tmpName As String
      Dim tmpList As New List(Of String)

      For Each kRow In WxTableRows
         For kCol As Integer = 0 To ParsedClassDefn(kRow).GetUpperBound(0)
            tmpName = ParsedClassDefn(kRow)(kCol).Trim.ToUpper
            If Not tmpName = "" Then
               If Not tmpList.Contains(tmpName) Then
                  tmpList.Add(tmpName)
               End If
            End If
         Next
      Next

      ReDim WxTable(tmpList.Count - 1)

      Try
         For kTable As Integer = 0 To tmpList.Count - 1
            Try
               SheetName = CStr(tmpList(kTable))
            Catch ex As Exception
               SheetName = ""
            End Try

            WxTable(kTable).Name = SheetName

            Dim vNode, vElement As String
            vNode = ""
            vElement = ""
            DivaCalc.ParseElementName(SheetName, vNode, vElement, vNodeName)
            IsMatrix = gv.NodeCollection.KeyExists(vNode) AndAlso (TypeOf gv.NodeCollection(vNode).CalcObject Is Matrix)

            If (Not IsMatrix) Then
               FireMsgBox("In node '" & vNodeName & "', mortality table '" & SheetName & "' is not defined as a matrix.")
               Return False
            End If
            If IsMatrix Then
               CType(gv.NodeCollection(vNode).CalcObject, Matrix).GetVectorSet(gVectorSet)
               t = Array.IndexOf(gVectorSet.ArrayNamesUC, vElement.ToUpper)
               If t >= 0 Then
                  With WxTable(kTable)

                     tmpArray = gVectorSet.ArrayValues(t)
                     .StartAge = CInt(tmpArray(0, 0))
                     .EndAge = CInt(tmpArray(tmpArray.GetUpperBound(0), 0))

                     ReDim .Wx(tmpArray.GetUpperBound(0) + .StartAge)
                     Dim tmpAge As Integer
                     For kRow = 0 To tmpArray.GetUpperBound(0)
                        tmpAge = CInt(tmpArray(kRow, 0))
                        .Wx(tmpAge) = tmpArray(kRow, 1)
                     Next kRow
                  End With

               Else
                  FireMsgBox("In node '" & vNodeName & "', the matrix named '" & vNode & "' does not contain an element named '" & vElement & "'.")
                  gv.MasterStop = True
                  Return False
               End If
            End If

         Next kTable

      Catch ex As Exception
         FireMsgBox("In node " & vNodeName & ", ReadWxSheets, " & ex.Message, MsgBoxStyle.Critical)
         gv.MasterStop = True
         Return False
      End Try
      Return True
   End Function

   Private Function ReadImprovSheets() As Boolean
      '   Read in a table and store results in Public Variable QxTable(Sex,Smk)
      Dim t As Integer
      Dim LineIn As String = ""
      'Dim ParsedIn() As String
      Dim kRow As Integer, SheetName As String
      'Dim UltCol, ThisAge, ParsedItems, kAge As Integer
      Dim FileIN As IO.StreamReader = Nothing
      Dim IsFile, IsMatrix As Boolean
      Dim gVectorSet As DivaCalcTools.Structures.VectorSet = Nothing
      Dim tmpArray(,) As Double

      'Get list of improvement table names
      Dim tmpName As String
      Dim tmpList As New List(Of String)

      For Each kRow In ImprovTableRows
         For kCol As Integer = 0 To ParsedClassDefn(kRow).GetUpperBound(0)
            tmpName = ParsedClassDefn(kRow)(kCol).Trim.ToUpper
            If Not tmpName = "" Then
               If Not tmpList.Contains(tmpName) Then
                  tmpList.Add(tmpName)
               End If
            End If
         Next
      Next

      ReDim ImprovTable(tmpList.Count - 1)

      Try
         For kTable As Integer = 0 To tmpList.Count - 1
            Try
               SheetName = tmpList(kTable)
            Catch ex As Exception
               SheetName = ""
            End Try

            ImprovTable(kTable).Name = SheetName
            'If SheetName = "Unity" Then
            '	'    "Unity" is key word - If so, set up table of 1's
            '	ImprovTable(kTable).EndAge = 106
            '	ImprovTable(kTable).SelectPer = 0
            '	ReDim ImprovTable(kTable).Qx(106, 0)
            '	For Index = 0 To 106
            '		ImprovTable(kTable).Qx(Index, 0) = 1.0
            '	Next Index
            'Else

            '   Read table from file
            '   Decide whether the table is a file or a Matrix Component
            Dim vNode, vElement As String
            IsFile = IO.File.Exists(SheetName)
            vNode = ""
            vElement = ""
            DivaCalc.ParseElementName(SheetName, vNode, vElement, vNodeName)
            IsMatrix = gv.NodeCollection.KeyExists(vNode) AndAlso (TypeOf gv.NodeCollection(vNode).CalcObject Is Matrix)
            If IsFile AndAlso IsMatrix Then
               FireMsgBox("In node '" & vNodeName & "', mortality table '" & SheetName & "' is both a matrix and a file name." _
                    & vbCr & "Matrix values assumed.")
               IsFile = False
            End If
            If (Not IsFile) AndAlso (Not IsMatrix) Then
               FireMsgBox("In node '" & vNodeName & "', mortality table '" & SheetName & "' is not defined as a matrix and is not a file name.")
               Return False
            End If
            If IsMatrix Then
               CType(gv.NodeCollection(vNode).CalcObject, Matrix).GetVectorSet(gVectorSet)
               t = Array.IndexOf(gVectorSet.ArrayNamesUC, vElement.ToUpper)
               If t >= 0 Then
                  With ImprovTable(kTable)
                     tmpArray = gVectorSet.ArrayValues(t)
                     .StartAge = CInt(tmpArray(1, 0))
                     .EndAge = CInt(tmpArray(tmpArray.GetUpperBound(0), 0))
                     '   copy in the Qx values, removing the row of select durations and the column of ages
                     .StartYear = CInt(tmpArray(0, 1))
                     .EndYear = CInt(tmpArray(0, tmpArray.GetUpperBound(1)))
                     ReDim .Rx(tmpArray.GetUpperBound(0) - 1 + .StartAge, tmpArray.GetUpperBound(1) - 1)
                     Dim tmpAge, kCol As Integer
                     For kRow = 1 To tmpArray.GetUpperBound(0)
                        tmpAge = CInt(tmpArray(kRow, 0))
                        For kCol = 1 To tmpArray.GetUpperBound(1)
                           .Rx(tmpAge, kCol - 1) = tmpArray(kRow, kCol)
                        Next kCol
                     Next kRow
                  End With

               Else
                  FireMsgBox("In node '" & vNodeName & "', the matrix named '" & vNode & "' does not contain an element named '" & vElement & "'.", MsgBoxStyle.Critical, gv.AppName)
                  gv.MasterStop = True
                  Return False
               End If
            End If

            'If IsFile Then
            '	FileIN = New IO.StreamReader(SheetName)
            '	For t = 1 To 200
            '		LineIn = FileIN.ReadLine
            '		If 0 < InStr(LineIn, "Minimum Age:", CompareMethod.Text) Then Exit For
            '	Next t
            '	ParsedIn = ParseToStr(LineIn)
            '	ImprovTable(kTable).StartAge = CInt(ParsedIn(2))

            '	LineIn = FileIN.ReadLine
            '	ParsedIn = ParseToStr(LineIn)
            '	ImprovTable(kTable).EndAge = CInt(ParsedIn(2))

            '	LineIn = FileIN.ReadLine
            '	ParsedIn = ParseToStr(LineIn)
            '	If 0 < InStr(LineIn, "Select Period", CompareMethod.Text) Then
            '		'   Table has select period and layout
            '		ParsedIn = ParseToStr(LineIn)
            '		ImprovTable(kTable).SelectPer = CInt(ParsedIn(2))
            '		UltCol = CInt(ParsedIn(2)) + 1
            '		LineIn = FileIN.ReadLine
            '		ParsedIn = ParseToStr(LineIn)
            '		ImprovTable(kTable).MaxSelectAge = CInt(ParsedIn(3))
            '		With ImprovTable(kTable)
            '			ReDim .Qx(.EndAge, .SelectPer)
            '		End With
            '	Else
            '		'   Table has ultimate period and different layout
            '		ImprovTable(kTable).SelectPer = 0
            '		UltCol = 0
            '		ImprovTable(kTable).MaxSelectAge = -1
            '		With QxTable(kTable)
            '			ReDim .Qx(.EndAge, 0)
            '		End With
            '	End If

            '	Do
            '		LineIn = FileIN.ReadLine
            '	Loop Until 0 < InStr(LineIn, "Table Values:", CompareMethod.Text)

            '	'  optional skip over Column Numbers
            '	If ImprovTable(kTable).MaxSelectAge > 0 Then LineIn = FileIN.ReadLine

            '	Do
            '		LineIn = FileIN.ReadLine
            '		ParsedIn = ParseToStr(LineIn)
            '		ParsedItems = ParsedIn.GetUpperBound(0)
            '		If Not IsNumeric(ParsedIn(0)) Then Exit Do
            '		ThisAge = CInt(ParsedIn(0))
            '		If ThisAge <= ImprovTable(kTable).MaxSelectAge Then
            '			If ThisAge > 115 Then Exit Do
            '			For t = 1 To ParsedItems - 1
            '				ImprovTable(kTable).Qx(ThisAge, t - 1) = CDbl(ParsedIn(t))
            '			Next t
            '		Else
            '			With ImprovTable(kTable)
            '				If ThisAge - .SelectPer > UBound(.Qx, 1) Then Exit Do
            '				.Qx(ThisAge - .SelectPer, UltCol) = CDbl(ParsedIn(1))
            '			End With
            '		End If
            '	Loop
            '	FileIN.Close()
            'End If
            'End If

         Next kTable

      Catch ex As Exception
         If Not IsNothing(FileIN) Then FileIN.Close()
         FireMsgBox("In node " & vNodeName & ", ReadImproveSheets, " & ex.Message, MsgBoxStyle.Critical)
         gv.MasterStop = True
         Return False
      End Try
      Return True

   End Function

   Private Function ReadQxSheets() As Boolean
      '   Read in a table and store results in Public Variable QxTable(Sex,Smk)
      Dim t As Integer
      Dim LineIn As String = ""
      Dim ParsedIn() As String
      Dim kRow As Integer, SheetName As String
      Dim Index, UltCol, ThisAge, ParsedItems, kAge As Integer
      Dim FileIN As IO.StreamReader = Nothing
      Dim IsFile, IsMatrix As Boolean
      Dim gVectorSet As DivaCalcTools.Structures.VectorSet = Nothing
      Dim tmpArray(,) As Double

      '  Read all classes to get list of table names
      Dim tmpName As String
      Dim tmpList As New List(Of String)

      For Each kRow In MortTableRows
         For kCol As Integer = 0 To ParsedClassDefn(kRow).GetUpperBound(0)
            tmpName = ParsedClassDefn(kRow)(kCol).Trim.ToUpper
            If tmpName = "" Then
               FireMsgBox("In node '" & vNodeName & "', Class definitions Row " & CStr(kRow + 1) & ", Column " & CStr(kCol + 1) _
                           & ", mortality table cannot be blank.")
               gv.MasterStop = True
               Return False
            End If
            If Not tmpList.Contains(tmpName) Then
               tmpList.Add(tmpName)
            End If
         Next
      Next

      ReDim QxTable(tmpList.Count - 1)

      Try
         For kTable As Integer = 0 To tmpList.Count - 1
            Try
               SheetName = CStr(tmpList(kTable))    'This used to be kRow. Is this correct?? 06/23/2016
            Catch ex As Exception
               SheetName = ""
            End Try

            QxTable(kTable).Name = SheetName
            If SheetName = "Unity" Then
               '    "Unity" is key word - If so, set up table of 1's
               QxTable(kTable).EndAge = 106
               QxTable(kTable).SelectPer = 0
               ReDim QxTable(kTable).Qx(106, 0)
               For Index = 0 To 106
                  QxTable(kTable).Qx(Index, 0) = 1.0
               Next Index
            Else
               '   Read table from file
               '   Decide whether the table is a file or a Matrix Component
               Dim vNode, vElement As String
               IsFile = IO.File.Exists(SheetName)
               vNode = ""
               vElement = ""
               DivaCalc.ParseElementName(SheetName, vNode, vElement, vNodeName)
               IsMatrix = gv.NodeCollection.KeyExists(vNode) AndAlso (TypeOf gv.NodeCollection(vNode).CalcObject Is Matrix)
               If IsFile AndAlso IsMatrix Then
                  FireMsgBox("In node '" & vNodeName & "', mortality table '" & SheetName & "' is both a matrix and a file name." _
                       & vbCr & "Matrix values assumed.")
                  IsFile = False
               End If
               If (Not IsFile) AndAlso (Not IsMatrix) Then
                  FireMsgBox("In node '" & vNodeName & "', mortality table '" & SheetName & "' is not defined as a matrix and is not a file name.")
                  Return False
               End If
               If IsMatrix Then
                  CType(gv.NodeCollection(vNode).CalcObject, Matrix).GetVectorSet(gVectorSet)
                  t = Array.IndexOf(gVectorSet.ArrayNamesUC, vElement.ToUpper)
                  If t >= 0 Then
                     With QxTable(kTable)
                        tmpArray = gVectorSet.ArrayValues(t)
                        .StartAge = CInt(tmpArray(1, 0))
                        .SelectPer = tmpArray.GetLength(1) - 2
                        .EndAge = CInt(tmpArray(tmpArray.GetUpperBound(0), 0))
                        For kAge = tmpArray.GetUpperBound(0) To 0 Step -1
                           If tmpArray(kAge, 1) > 0.0 Then
                              .MaxSelectAge = CInt(tmpArray(kAge, 0))
                              Exit For
                           End If
                        Next kAge
                        '   copy in the Qx values, removing the row of select durations and the column of ages
                        ReDim .Qx(tmpArray.GetUpperBound(0) - 1 + .StartAge, tmpArray.GetUpperBound(1) - 1)
                        Dim tmpAge, kCol As Integer
                        For kRow = 1 To tmpArray.GetUpperBound(0)
                           tmpAge = CInt(tmpArray(kRow, 0))
                           For kCol = 1 To tmpArray.GetUpperBound(1)
                              .Qx(tmpAge, kCol - 1) = tmpArray(kRow, kCol)
                           Next kCol
                        Next kRow
                     End With

                  Else
                     FireMsgBox("In node '" & vNodeName & "', the matrix named '" & vNode & "' does not contain an element named '" & vElement & "'.", MsgBoxStyle.Critical, gv.AppName)
                     gv.MasterStop = True
                     Return False
                  End If
               End If

               If IsFile Then
                  FileIN = New IO.StreamReader(SheetName)
                  For t = 1 To 200
                     LineIn = FileIN.ReadLine
                     If 0 < InStr(LineIn, "Minimum Age:", CompareMethod.Text) Then Exit For
                  Next t
                  ParsedIn = ParseToStr(LineIn)
                  QxTable(kTable).StartAge = CInt(ParsedIn(2))

                  LineIn = FileIN.ReadLine
                  ParsedIn = ParseToStr(LineIn)
                  QxTable(kTable).EndAge = CInt(ParsedIn(2))

                  LineIn = FileIN.ReadLine
                  ParsedIn = ParseToStr(LineIn)
                  If 0 < InStr(LineIn, "Select Period", CompareMethod.Text) Then
                     '   Table has select period and layout
                     ParsedIn = ParseToStr(LineIn)
                     QxTable(kTable).SelectPer = CInt(ParsedIn(2))
                     UltCol = CInt(ParsedIn(2)) + 1
                     LineIn = FileIN.ReadLine
                     ParsedIn = ParseToStr(LineIn)
                     QxTable(kTable).MaxSelectAge = CInt(ParsedIn(3))
                     With QxTable(kTable)
                        ReDim .Qx(.EndAge, .SelectPer)
                     End With
                  Else
                     '   Table has ultimate period and different layout
                     QxTable(kTable).SelectPer = 0
                     UltCol = 0
                     QxTable(kTable).MaxSelectAge = -1
                     With QxTable(kTable)
                        ReDim .Qx(.EndAge, 0)
                     End With
                  End If

                  Do
                     LineIn = FileIN.ReadLine
                  Loop Until 0 < InStr(LineIn, "Table Values:", CompareMethod.Text)

                  '  optional skip over Column Numbers
                  If QxTable(kTable).MaxSelectAge > 0 Then LineIn = FileIN.ReadLine

                  Do
                     LineIn = FileIN.ReadLine
                     ParsedIn = ParseToStr(LineIn)
                     ParsedItems = ParsedIn.GetUpperBound(0)
                     If Not IsNumeric(ParsedIn(0)) Then Exit Do
                     ThisAge = CInt(ParsedIn(0))
                     If ThisAge <= QxTable(kTable).MaxSelectAge Then
                        If ThisAge > 115 Then Exit Do
                        For t = 1 To ParsedItems - 1
                           QxTable(kTable).Qx(ThisAge, t - 1) = CDbl(ParsedIn(t))
                        Next t
                     Else
                        With QxTable(kTable)
                           If ThisAge - .SelectPer > UBound(.Qx, 1) Then Exit Do
                           .Qx(ThisAge - .SelectPer, UltCol) = CDbl(ParsedIn(1))
                        End With
                     End If
                  Loop
                  FileIN.Close()
               End If
            End If

         Next kTable

      Catch ex As Exception
         If Not IsNothing(FileIN) Then FileIN.Close()
         FireMsgBox("In node " & vNodeName & ", ReadQxSheets, " & ex.Message, MsgBoxStyle.Critical)
         gv.MasterStop = True
         Return False
      End Try
      Return True
   End Function

   Private Function ComputeIxFactors(ByVal ImprovTable As ImprovementTable, ByVal MortTableYear As Integer, ByVal numYears As Integer) As ImprovementFactors
      Dim CompoundImprov As Double
      Dim outputFactors As ImprovementFactors

      outputFactors.BaseYear = MortTableYear
      outputFactors.numYears = numYears
      outputFactors.StartAge = ImprovTable.StartAge
      outputFactors.EndAge = ImprovTable.EndAge

      ReDim outputFactors.Ix(ImprovTable.EndAge, numYears - 1)

      For kAge As Integer = ImprovTable.StartAge To ImprovTable.EndAge
         CompoundImprov = 1.0
         For kYear As Integer = 0 To numYears - 1
            'If kYear = 0, we keep compound improvement factor at 1 (so do nothing)
            '--> b/c year zero is base year. So no improvement
            If kYear > 0 Then
               'If kYear > 0, then apply following logic:
               If kYear <= ImprovTable.StartYear - 1 - outputFactors.BaseYear Then
                  'This means you're not yet at the mortality improvement start year
                  'E.g. Base Year 2014, Improvement Start Year 2016, then for kYear <= 1 = 2016-2014-1, improvement has not started 
                  'Keep compound improvement factor at 1
               ElseIf kYear <= ImprovTable.EndYear - outputFactors.BaseYear Then
                  'This means you're in the improvement years, but you haven't reached the end of the table yet.
                  CompoundImprov *= 1.0 - ImprovTable.Rx(kAge, kYear - (ImprovTable.StartYear) + outputFactors.BaseYear)     'Multiply by correct rate
               Else
                  'This means you're past the end of the table, so just apply final year's rate
                  CompoundImprov *= 1.0 - ImprovTable.Rx(kAge, ImprovTable.EndYear - ImprovTable.StartYear) 'Multiply by last rate of table
               End If
            End If
            outputFactors.Ix(kAge, kYear) = CompoundImprov
         Next
      Next
      Return outputFactors
   End Function

   Function ParseToStr(ByVal InString As String) As String()
      '   Parse based on blanks, but ignore leading, trailing and multiple occurences
      '   Treat Tab as a single blank
      InString = Replace(InString, vbTab, " ")
      InString = InString.Trim
      Do While InString.IndexOf("  ") >= 0
         InString = Replace(InString, "  ", " ")
      Loop
      ParseToStr = Split(InString, " ")
   End Function

   Public Overrides Sub ConvertDateInputs(ByRef ParamStr As String, ByRef AnyChanges As Boolean, ByRef MonthList As String())
      '   Convert dates from one region to another
      Dim vParam(), vRows(), vCols() As String
      Dim kRow, kCol As Integer
      Dim DateCols() As Integer = {1, 4, 11, 14}   '    Date columns
      vParam = Split(ParamStr, vbFormFeed)
      vRows = Split(vParam(sTabSeriatimIn), vbCr)
      AnyChanges = False
      For kRow = 0 To vRows.GetUpperBound(0)
         vCols = Split(vRows(kRow), vbTab)
         For Each kCol In DateCols
            Try
               If vCols(kCol).Trim <> "" AndAlso Not IsNumeric(vCols(kCol)) Then '    Assume an all-numeric content is a defined format
                  vCols(kCol) = DivaCalc.RegionalConvertDate(vCols(kCol), MonthList)
                  AnyChanges = True
               End If
            Catch ex As Exception
               FireMsgBox("In date conversions for '" & ParentNode.NodeName & "', invalid dates: " & vbCr & ex.Message)
            End Try
         Next kCol
         If AnyChanges Then vRows(kRow) = Join(vCols, vbTab)
      Next kRow
      If AnyChanges Then
         vParam(sTabSeriatimIn) = Join(vRows, vbCr)
         ParamStr = Join(vParam, vbFormFeed)
      Else
         '   Nothing
      End If
   End Sub

#End Region

#Region "Usual Overrides"

   Public Overrides Sub SetGlobalCalcFunctions(ByRef CalcFunc As DivaCalcTools.CalcFunctions, _
         ByRef Xgv As DivaCalcTools.GlobalVariables)
      DivaCalc = CalcFunc
      gv = Xgv
   End Sub

   Overrides ReadOnly Property CalcName() As String
      Get
         CalcName = "ActFuncB"
      End Get
   End Property

   Overrides Property ElementNames() As String()
      Get
         ElementNames = vEleNames
      End Get
      Set(ByVal Value As String())
         '   nothing
      End Set
   End Property

   Overrides ReadOnly Property ElementArray() As Double(,)
      Get
         ElementArray = vEleVal
      End Get
   End Property

   Overrides ReadOnly Property ElementIndex(ByVal ElementName As String) As Integer
      Get
         For ElementIndex = 0 To sMaxElements
            If ElementName = vEleNames(ElementIndex) Then Exit Property
         Next ElementIndex

         ElementIndex = -1
         gv.MasterStop = True
      End Get
   End Property

   Overrides ReadOnly Property ValueElementTime(ByVal ElementIndex As Integer, ByVal TimePeriod As Integer) As Double
      Get
         If TimePeriod >= 0 Then
            ValueElementTime = vEleVal(ElementIndex, TimePeriod)
         Else
            ValueElementTime = 0.0
         End If
      End Get
   End Property

   Overrides Property ParameterString() As String
      Get
         ParameterString = Join(vParamArray, vbFormFeed)
      End Get
      Set(ByVal Value As String)
         vParamArray = Split(Value, vbFormFeed)
         If vParamArray.GetUpperBound(0) < 3 Then
            ReDim Preserve vParamArray(3)
            vParamArray(3) = ""
         End If
      End Set
   End Property

   Overrides ReadOnly Property ParamNames() As String()
      Get
         Select Case vPropertyTabIndex
            Case sTabParameters
               Return vParamNames
            Case sTabClasses
               Return vClassItems
            Case Else
               ParamNames = Nothing
         End Select
      End Get
   End Property

   Overrides ReadOnly Property ParamTypes() As Integer()
      Get
         Dim gT As Integer = DivaCalcTools.GridConstants.sParamTypeText
         Dim g11 As Integer = DivaCalcTools.GridConstants.sParamTypeList11
         Dim g13 As Integer = DivaCalcTools.GridConstants.sParamTypeList13
         Dim g14 As Integer = DivaCalcTools.GridConstants.sParamTypeList14
         Dim g17 As Integer = DivaCalcTools.GridConstants.sParamTypeList17
         Dim g19 As Integer = DivaCalcTools.GridConstants.sParamTypeList19
         Dim g24 As Integer = DivaCalcTools.GridConstants.sParamTypeList24
         Dim gYN As Integer = DivaCalcTools.GridConstants.sParamTypeYesNo
         Select Case vPropertyTabIndex
            Case sTabParameters
               ParamTypes = {gT, DivaCalcTools.GridConstants.sParamTypeList16, gT, _
                   g24, gT, DivaCalcTools.GridConstants.sParamTypeList20, _
                   DivaCalcTools.GridConstants.sParamTypeList15, _
                   DivaCalcTools.GridConstants.sParamTypeList12, DivaCalcTools.GridConstants.sParamTypeList12, _
                   gT, g17, gT, gT, gT, _
                   DivaCalcTools.GridConstants.sParamTypeList22, gT, DivaCalcTools.GridConstants.sParamTypeList23, _
                             gYN}
            Case sTabClasses
               ParamTypes = {gT, gT, gT, gT, gT, gT, gT, gT, gT, gT, gT, DivaCalcTools.GridConstants.sParamTypeList21, gT, gT, _
                  gYN, gT, gYN, gT, gT, _
                  g13, gT, gT, g19, gT, gT, _
                  gT, gT, gT, gT, gT, gT, gT, gT, gT, gT, DivaCalcTools.GridConstants.sParamTypeList21, gT, gT, _
                  gYN, gT, gYN, gT, gT, _
                  g13, gT, gT, g19, gT, gT, _
                  gT, gT, gT, gT, gT, gT, gT, gT, gT, gT, DivaCalcTools.GridConstants.sParamTypeList21, gT, gT, _
                  gYN, gT, gYN, gT, gT, _
                  g13, gT, gT, g19, gT, gT}
            Case Else
               ParamTypes = Nothing
         End Select
      End Get
   End Property

   Public Overrides ReadOnly Property ProtectedRows() As Integer()
      Get
         Select Case Me.vPropertyTabIndex
            Case sTabParameters
               ProtectedRows = {10000 + kRowsParams.HdrMortBasis, 10000 + kRowsParams.HdrSeriatim, 10000 + kRowsParams.HdrBenefits, _
                  10000 + kRowsParams.HdrOutputs}
            Case sTabClasses
               ProtectedRows = {10000 + kRowsClasses.HdrMortBasis_GC, 10000 + kRowsClasses.HdrBenefits_GC, _
                            10000 + kRowsClasses.HdrMortBasis_SY, 10000 + kRowsClasses.HdrBenefits_SY, _
                            10000 + kRowsClasses.HdrMortBasis_OT, 10000 + kRowsClasses.HdrBenefits_OT}
            Case Else
               ProtectedRows = Nothing
         End Select
      End Get
   End Property

   Overrides ReadOnly Property HasElements() As Boolean
      Get
         HasElements = True
      End Get
   End Property

   'Public Overrides ReadOnly Property ColmnIsTimeSeries() As Boolean
   '    Get
   '        ColmnIsTimeSeries = False
   '    End Get
   'End Property

   Overrides ReadOnly Property ColumnHeadings() As String()
      Get
         Select Case vPropertyTabIndex
            Case sTabParameters
               Return {"", "Value"}
            Case sTabClasses
               Dim kList() As String
               Dim k, kCols As Integer
               kCols = Split(Split(vParamArray(sTabClasses), vbCr)(0), vbTab).GetUpperBound(0) + 1
               ReDim kList(kCols)

               kList(0) = ""
               For k = 1 To kCols
                  kList(k) = "Class " & CStr(k) & ": " & Split(Split(vParamArray(sTabClasses), vbCr)(kRowsClasses.className), vbTab)(k - 1)
               Next
               Return kList

               'to do
               ' Return {"Class 1", "Class 2"}
            Case sTabSeriatimIn
               Return CType(vColSeriatimIn.Clone, String())
            Case sTabSeriatimOut
               If HasAnySeriatimOut Then
                  Return Split(OutColHeading, vbTab)
               Else
                  Return {"No output"}
               End If
            Case Else
               Return Nothing
         End Select
      End Get
   End Property


   Public Overrides ReadOnly Property ImpliedInNodes() As System.Collections.Specialized.StringCollection
      Get
         '            ImpliedInNodes = DivaCalc.fnImpliedInNodes(CellType(0), ParmNode(0), FuncExpressions(0), gv.NodeCollection)
         Dim kParm As Integer, SubList As Specialized.StringCollection, kItem As String
         ImpliedInNodes = New Specialized.StringCollection
         For kParm = 0 To ParmNode.GetUpperBound(0)
            SubList = DivaCalc.fnImpliedInNodes(CellType(kParm), ParmNode(kParm), FuncExpressions(kParm), gv.NodeCollection)
            For Each kItem In SubList
               If Not ImpliedInNodes.Contains(kItem) Then ImpliedInNodes.Add(kItem)
            Next
         Next kParm
      End Get
   End Property

   Public Overrides ReadOnly Property CheckCirculars As String
      Get
         Return DivaCalc.fnCheckCircularsTabbed(CellType, ParmNode, FuncExpressions, Me.CalcOrder)
      End Get
   End Property

   Public Overrides ReadOnly Property HasPropertyTabs() As Boolean
      Get
         HasPropertyTabs = True
      End Get
   End Property

   Public Overrides ReadOnly Property PropertyTabText() As String()
      Get
         PropertyTabText = vPropertyTabText
      End Get
   End Property

   Public Overrides Property PropertyTabIndex() As Integer
      Get
         PropertyTabIndex = vPropertyTabIndex
      End Get
      Set(ByVal Value As Integer)
         vPropertyTabIndex = Value
      End Set
   End Property

   Public Overrides Sub SetList(ByRef Index As Integer, ByRef MyList() As String)
      Select Case vPropertyTabIndex
         Case sTabParameters, sTabClasses
            Select Case Index
               Case DivaCalcTools.GridConstants.sParamTypeList11
                  MyList = CType(vIntrMethods.Clone, String())
               Case DivaCalcTools.GridConstants.sParamTypeList12
                  MyList = {"Annual", "Semi-Annual", "Quarterly", "Monthly", "Semi-Monthly", "Weekly", "Daily"}
               Case DivaCalcTools.GridConstants.sParamTypeList13
                  MyList = {"Rate", "Vector"}
               Case DivaCalcTools.GridConstants.sParamTypeList14
                  MyList = {"Age", "Seriatim Date", "Seriatim Age"}
               Case DivaCalcTools.GridConstants.sParamTypeList15
                  MyList = {"No Change", "Annuity Changes"}
               Case DivaCalcTools.GridConstants.sParamTypeList16
                  MyList = {"Age Last", "Age Nearest", "Interpolate", "Interpolate Months"}
               Case DivaCalcTools.GridConstants.sParamTypeList17
                  MyList = CType(CashFlowTiming.Clone, String())
               Case DivaCalcTools.GridConstants.sParamTypeList18
                  MyList = {"Dummy Item A", "Dummy Item B"}
               Case DivaCalcTools.GridConstants.sParamTypeList19
                  MyList = CType(vInflMethods.Clone, String())
               Case DivaCalcTools.GridConstants.sParamTypeList20
                  ReDim MyList(1)
                  MyList(0) = "None"
                  'MyList(1) = "Seriatim Date"
                  MyList(1) = "Seriatim Years"
                  'Dim k As Integer
                  'For k = 1 To 30
                  '	MyList(2 + k) = CStr(k)
                  'Next k
               Case DivaCalcTools.GridConstants.sParamTypeList21
                  MyList = {"None", "Inputs", "Seriatim", "Seriatim [M1 * Qx + M2]"}
               Case DivaCalcTools.GridConstants.sParamTypeList22
                  MyList = {"None", "Initial TimePeriod, All Scenarios", "All TimePeriods, First Scenario", "All TimePeriods, All Scenarios"}
               Case DivaCalcTools.GridConstants.sParamTypeList23
                  MyList = {"Exponential", "Linear"}
               Case DivaCalcTools.GridConstants.sParamTypeList24
                  MyList = {"Age", "Date"}
            End Select
         Case Else
      End Select
   End Sub

   Public Overrides ReadOnly Property InputGridSetting() As DivaCalcTools.Structures.InputGridSettings
      Get
         Dim kInputSetting As New DivaCalcTools.Structures.InputGridSettings
         With kInputSetting
            .ColIsTimeSeries = False
            .IsTriangle = False
            .StoreRows = 1
            Select Case vPropertyTabIndex
               Case sTabParameters
                  .AllowColumnSort = False
                  .FixedCols = 1
                  .FixedRows = 1
                  .AllowAdjustCols = False
                  .AllowAdjustRows = False
                  .StoreCols = 1
               Case sTabClasses
                  .AllowColumnSort = False
                  .FixedCols = 1
                  .FixedRows = 1
                  .AllowAdjustCols = True
                  .AllowAdjustRows = False
                  .StoreCols = 1

               Case sTabSeriatimIn
                  .AllowColumnSort = True
                  .FixedCols = 0
                  .FixedRows = 1
                  .AllowAdjustCols = False
                  .AllowAdjustRows = True
                  .StoreCols = 0
               Case sTabSeriatimOut
                  .AllowColumnSort = True
                  .FixedCols = 0
                  .FixedRows = 1
                  .AllowAdjustCols = False
                  .AllowAdjustRows = True
                  .StoreCols = 0

            End Select
         End With
         InputGridSetting = kInputSetting
      End Get
   End Property

   Public Overrides Property ParameterStringTab() As String
      Get
         ParameterStringTab = vParamArray(vPropertyTabIndex)
      End Get
      Set(ByVal Value As String)
         Try
            vParamArray(vPropertyTabIndex) = Value
         Catch ex As Exception
            FireMsgBox("In node " & vNodeName & vbCrLf & ex.Message, MsgBoxStyle.Critical)
         End Try
      End Set
   End Property

   Public Overrides Property ElementDescriptions() As String()
      Get
         ElementDescriptions = vEleDescriptions
      End Get
      Set(ByVal Value As String())

      End Set
   End Property

   Public Overrides Sub ComboBoxChangedAction(ByVal Row As Integer, ByVal Col As Integer, ByVal SelectedIndex As Integer, ByVal ComboText As String, _
         ByRef ItemsToEnable() As Integer, ByRef ItemsToDisable() As Integer, ByRef Action As String)
      Select Case vPropertyTabIndex
         Case sTabParameters
            Action = "Rows"
            Select Case Row - 1
               Case kRowsParams.SupplementaryOut
                  Select Case ComboText
                     Case "None"
                        ItemsToDisable = {1 + kRowsParams.SupFileName}
                        ItemsToEnable = Nothing
                     Case Else
                        ItemsToDisable = Nothing
                        ItemsToEnable = {1 + kRowsParams.SupFileName}
                  End Select
            End Select
         Case sTabClasses
            Action = "Cells"

            Select Case Split(vParamArray(sTabParameters), vbCr)(kRowsParams.UseRetireLoop)
               Case "Yes"
                  Action = "Rows"
                  ItemsToEnable = {kRowsClasses.RetLoopAgeFactors_GC + 1, kRowsClasses.RetLoopAgeFactors_SY + 1, kRowsClasses.RetLoopAgeFactors_OT + 1}
               Case "No"
                  Action = "Rows"
                  ItemsToDisable = {kRowsClasses.RetLoopAgeFactors_GC + 1, kRowsClasses.RetLoopAgeFactors_SY + 1, kRowsClasses.RetLoopAgeFactors_OT + 1}
            End Select

            Select Case Row - 1
               Case kRowsClasses.MortPctType_GC
                  Action = "Cells"
                  Select Case ComboText
                     Case "None", "Seriatim", "Seriatim [M1 * Qx + M2]"
                        ItemsToEnable = Nothing
                        ItemsToDisable = {1 + kRowsClasses.MortPctMale_GC, Col, 1 + kRowsClasses.MortPctFem_GC, Col}
                     Case "Inputs"
                        ItemsToEnable = {1 + kRowsClasses.MortPctMale_GC, Col, 1 + kRowsClasses.MortPctFem_GC, Col}
                        ItemsToDisable = Nothing
                  End Select
               Case kRowsClasses.MortPctType_SY
                  Action = "Cells"
                  Select Case ComboText
                     Case "None", "Seriatim", "Seriatim [M1 * Qx + M2]"
                        ItemsToEnable = Nothing
                        ItemsToDisable = {1 + kRowsClasses.MortPctMale_SY, Col, 1 + kRowsClasses.MortPctFem_SY, Col}
                     Case "Inputs"
                        ItemsToEnable = {1 + kRowsClasses.MortPctMale_SY, Col, 1 + kRowsClasses.MortPctFem_SY, Col}
                        ItemsToDisable = Nothing
                  End Select
               Case kRowsClasses.MortPctType_OT
                  Action = "Cells"
                  Select Case ComboText
                     Case "None", "Seriatim", "Seriatim [M1 * Qx + M2]"
                        ItemsToEnable = Nothing
                        ItemsToDisable = {1 + kRowsClasses.MortPctMale_OT, Col, 1 + kRowsClasses.MortPctFem_OT, Col}
                     Case "Inputs"
                        ItemsToEnable = {1 + kRowsClasses.MortPctMale_OT, Col, 1 + kRowsClasses.MortPctFem_OT, Col}
                        ItemsToDisable = Nothing
                  End Select
               Case kRowsClasses.UseUnisex_GC
                  Action = "Cells"
                  Select Case ComboText
                     Case "Yes"
                        ItemsToEnable = {1 + kRowsClasses.UnisexMalePct_GC, Col}
                        ItemsToDisable = Nothing
                     Case "No"
                        ItemsToDisable = {1 + kRowsClasses.UnisexMalePct_GC, Col}
                        ItemsToEnable = Nothing
                  End Select
               Case kRowsClasses.UseUnisex_SY
                  Action = "Cells"
                  Select Case ComboText
                     Case "Yes"
                        ItemsToEnable = {1 + kRowsClasses.UnisexMalePct_SY, Col}
                        ItemsToDisable = Nothing
                     Case "No"
                        ItemsToDisable = {1 + kRowsClasses.UnisexMalePct_SY, Col}
                        ItemsToEnable = Nothing
                  End Select
               Case kRowsClasses.UseUnisex_OT
                  Action = "Cells"
                  Select Case ComboText
                     Case "Yes"
                        ItemsToEnable = {1 + kRowsClasses.UnisexMalePct_OT, Col}
                        ItemsToDisable = Nothing
                     Case "No"
                        ItemsToDisable = {1 + kRowsClasses.UnisexMalePct_OT, Col}
                        ItemsToEnable = Nothing
                  End Select
               Case kRowsClasses.IntrRateOrVec_GC, kRowsClasses.InflRateOrVec_GC, _
                     kRowsClasses.IntrRateOrVec_SY, kRowsClasses.InflRateOrVec_SY, _
                     kRowsClasses.IntrRateOrVec_OT, kRowsClasses.InflRateOrVec_OT
                  Action = "Cells"
                  Select Case ComboText
                     Case "Rate"
                        ItemsToEnable = {Row + 1, Col}
                        ItemsToDisable = {Row + 2, Col}
                     Case "Vector"
                        ItemsToEnable = {Row + 2, Col}
                        ItemsToDisable = {Row + 1, Col}
                     Case "Seriatim", "Seriatim Vector"
                        ItemsToEnable = Nothing
                        ItemsToDisable = {Row + 1, Col, Row + 2, Col}
                  End Select
            End Select
         Case Else
      End Select
   End Sub

   Public Overrides ReadOnly Property CommandRows() As Integer()
      Get
         Select Case vPropertyTabIndex
            Case sTabParameters
               CommandRows = {1 + kRowsParams.SeriatimOutCol, 1 + kRowsParams.SeriatimFmt}
            Case Else
               CommandRows = Nothing
         End Select
      End Get
   End Property

   Public Overrides Sub CommandButtonFired(ByRef Row As Integer, ByRef Col As Integer, ByRef CellContent(,) As String)
      Select Case Row - 1
         Case kRowsParams.SeriatimOutCol
            Try
               Dim fPick As New netToolKit.frmPickList
               Dim k As Integer, lItem As Windows.Forms.ListViewItem
               fPick.Text = vNodeName & " - Pick output columns"
               fPick.Label1.Font = gv.DefaultFont
               fPick.LabelText = "Items can be re-ordered by dragging up or down list."
               With fPick.lstAgg
                  .Font = gv.DefaultFont
                  .View = Windows.Forms.View.Details
                  .CheckBoxes = True
                  .SuspendLayout()
                  .Columns.Add("Item", 100, Windows.Forms.HorizontalAlignment.Right)
                  .Columns.Add("Label", 300, Windows.Forms.HorizontalAlignment.Left)
                  '   Build the list in the order given then add remaining items unchecked
                  Dim tmpList As String = CellContent(Row, Col)
                  Dim tmpSpl() As String
                  Do While 0 <= tmpList.IndexOf("  ")
                     tmpList = Replace(tmpList, "  ", " ")
                  Loop
                  tmpSpl = Split(tmpList.Trim, " ")
                  For k = 0 To tmpSpl.GetUpperBound(0)
                     Try
                        If tmpSpl(k).Trim <> "" Then
                           lItem = .Items.Add(tmpSpl(k))
                           lItem.SubItems.Add(vColSeriatimOut(CInt(tmpSpl(k))))
                           lItem.Checked = True
                        End If
                     Catch ex As Exception
                        '   Ignore - some faulty input
                     End Try
                  Next k
                  For k = 0 To vColSeriatimOut.GetUpperBound(0)
                     If Array.IndexOf(tmpSpl, Format(k, "00")) < 0 Then
                        lItem = .Items.Add(Format(k, "00"))
                        lItem.SubItems.Add(vColSeriatimOut(k))
                        lItem.Checked = False
                     End If
                  Next k
                  .ResumeLayout()
               End With
               fPick.AllowItemReorder = True
               fPick.AllowSort = True
               fPick.SortOnOKclick = False
               If fPick.ShowDialog = Windows.Forms.DialogResult.OK Then
                  If IsNothing(fPick.CheckedItems) Then
                     CellContent(Row, Col) = ""
                  Else
                     CellContent(Row, Col) = fPick.CheckedItems
                  End If
               End If
            Catch ex As Exception
               FireMsgBox("In Diva ActFuncs " & vNodeName & ":" & vbCrLf & ex.Message)
            End Try
         Case kRowsParams.SeriatimFmt
            Dim frmF As New netToolKit.frmNumberFmt
            With frmF
               .EnableColumns = False
               .EnableColWids = False
               .EnableAlignment = False
               .FormatCode = CellContent(Row, Col)
               If .ShowDialog = Windows.Forms.DialogResult.OK Then
                  CellContent(Row, Col) = .FormatCode
               End If
            End With
      End Select
   End Sub

   Function ExtendVector(ByVal vecToExtend As Double(), newLength As Integer) As Double()
      'Input: Vector and Desired Length of new vector
      'Output: New Vector with desired length. Values are equal to the input vector, with the last value repeated
      'If the desired length is less than or equal to the original length, then the function simply returns the old vector (i.e. doesn't truncate)
      If newLength <= vecToExtend.Length Then Return vecToExtend

      Dim newVec(newLength - 1) As Double
      For k As Integer = 0 To newLength - 1
         If k <= vecToExtend.Length - 1 Then
            newVec(k) = vecToExtend(k)
         Else
            newVec(k) = vecToExtend(vecToExtend.Length - 1)
         End If
      Next

      Return newVec

   End Function

   Function InterpretRetireAgeFactors(ByVal inputString As String) As Double(,)
      Dim trimString As String = inputString
      Dim tmpVecStr As String()
      Dim parsedAgeFactors As Double(,)

      trimString = inputString.Trim
      Do While trimString.IndexOf("  ") >= 0
         trimString = trimString.Replace("  ", " ")
      Loop

      'trimString = System.Text.RegularExpressions.Regex.Replace(inputString, "[ ]{2,}", " ")
      tmpVecStr = Split(trimString, " ")

      If tmpVecStr.GetUpperBound(0) Mod 2 = 0 Then
         Err.Raise(1001, "Retirement Loop Age Factors", "Number of retirement ages does not equal number of factors")
      End If

      ReDim parsedAgeFactors(CInt((tmpVecStr.GetUpperBound(0) - 1) / 2), 1)

      For k As Integer = 0 To tmpVecStr.GetUpperBound(0) Step 2
         parsedAgeFactors(k \ 2, 0) = CDbl(tmpVecStr(k))
         parsedAgeFactors(k \ 2, 1) = CDbl(tmpVecStr(k + 1))
      Next

      Return parsedAgeFactors

   End Function

   Function InterpolateRetireAgeFactors(ByVal inputArray As Double(,)) As Double(,)
      Dim startAgeYrs As Double = Math.Round(inputArray(0, 0) * 12) / 12                        'First age to run, in years, to nearest month
      Dim endAgeYrs As Double = Math.Round(inputArray(inputArray.GetUpperBound(0), 0) * 12) / 12      'Last age to run, in years, to nearest month

      If startAgeYrs > endAgeYrs Then
         Throw New Exception("When running multiple retirement ages, ages must be distinct and in sequential order.")
      End If

      Dim numAgesToRun As Integer = CInt((endAgeYrs - startAgeYrs) * 12) + 1
      Dim outputArray(numAgesToRun - 1, 1) As Double

      Dim AgeBelow, AgeAbove, indAgeBelow As Integer
      Dim FactorBelow, FactorAbove, interpLength, interpFactor As Double

      Dim currAgeMth As Integer = CInt(startAgeYrs * 12)

      indAgeBelow = 0                                 'Index of age below
      AgeBelow = CInt(Math.Round(inputArray(indAgeBelow, 0) * 12))              'Age below (in months)
      If numAgesToRun > 1 Then
         AgeAbove = CInt(Math.Round(inputArray(indAgeBelow + 1, 0) * 12))          'Age above (in months)
      Else
         AgeAbove = -1  'Not needed, so shouldn't produce an error...
      End If

      FactorBelow = inputArray(indAgeBelow, 1)        'Factor Below
      If numAgesToRun > 1 Then
         FactorAbove = inputArray(indAgeBelow + 1, 1)    'Factor Above
      Else
         FactorAbove = -1
      End If

      interpLength = AgeAbove - AgeBelow        'Length of interpolation interval, in months

      If numAgesToRun > 1 Then
         For kAgeMonth As Integer = 0 To numAgesToRun - 2 'Maximum age handled outside loop
            If currAgeMth = AgeAbove Then
               AgeBelow = AgeAbove
               indAgeBelow += 1
               AgeAbove = CInt(Math.Round(inputArray(indAgeBelow + 1, 0) * 12))
               FactorBelow = inputArray(indAgeBelow, 1)
               FactorAbove = inputArray(indAgeBelow + 1, 1)
               interpLength = AgeAbove - AgeBelow
            End If

            'If the age above is less than or equal to the age below, throw an exception.
            If AgeBelow >= AgeAbove Then
               Throw New Exception("Error in " & vNodeName & ": When running multiple retirement ages, ages must be distinct and in sequential order.")
            End If

            interpFactor = (currAgeMth - AgeBelow) / interpLength
            outputArray(kAgeMonth, 0) = currAgeMth / 12
            outputArray(kAgeMonth, 1) = interpFactor * FactorAbove + (1 - interpFactor) * FactorBelow
            currAgeMth += 1
         Next
      End If
      outputArray(numAgesToRun - 1, 0) = inputArray(inputArray.GetUpperBound(0), 0)
      outputArray(numAgesToRun - 1, 1) = inputArray(inputArray.GetUpperBound(0), 1)

      Return outputArray
   End Function

#End Region

End Class