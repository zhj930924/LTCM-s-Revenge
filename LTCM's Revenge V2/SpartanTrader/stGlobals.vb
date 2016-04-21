Module stGlobals
    Public ArbYesterday As Boolean = False


    Public teamID As String = "16"
    Public portfolioTableName As String = "PortfolioTeam" + teamID
    Public ConfirmationTicketTableName As String = "ConfirmationTicketTeam" + teamID

    Public RecArray(12) As Recommendation  ' index goes from 0 - 11


    Public traderMode As String = "Manual"
    Public secondsLeft As Integer = 0

    'HM18
    Public excludeIP As Boolean = False
    Public TPVNoHedge As Double = 0

    'HM17
    Public CAccountAT As Double = 0
    Public marginAT As Double = 0

    'HM16
    Public myTransaction As Transaction = New Transaction

    'HM15
    Public CAccount As Double = 0
    Public margin As Double = 0
    Public AP As Double = 0
    Public TPV As Double = 0
    Public TaTPV As Double = 0
    Public TE As Double = 0
    Public lastTransactionDate As Date
    Public interestSLT As Double = 0
    Public TEpercent As Double = 0
    Public lastPriceDownloadDate As Date

    'HM14

    Public initialCAccount As Double = 0
    Public iRate As Double = 0
    Public startDate As Date
    Public currentDate As Date
    Public endDate As Date
    Public maxMargins As Double = 0
    Public TPVatStart As Double = 0
    Public IP As Double = 0

    'HM13
    Public activeDB As String = ""
End Module
