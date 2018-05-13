Attribute VB_Name = "Enums"
Enum TaskPriority
    NoPriority = 1
    LowPriority = 2
    MediumPriority = 3
    HighPriority = 4
End Enum

Enum TaskPurpose
    None = 0
    MembershipSales = 7
    InternalSales = 8
    Retention = 9
    ProductSales = 10
    OtherSales = 11
    Collections = 20
    Admin = 87
End Enum

Enum TaskColor
    White = 0
    Red = 1
    Green = 2
    Blue = 3
    oRange = 4
    Purple = 5
End Enum
    

Enum icdCollections
    IncludeCollections = 1004
    DontIncludeCollections = 1
End Enum

Enum icdPayType
    AnyPayType = 1004
    CCOnly = 1
    ACHOnly = 2
    NoPayType = 3
End Enum

Enum icdInclusion
    Include = 1004
    DontInclude = 1
End Enum

Enum icdAllPastInvoices
    IncludeAll = True
    DontIncludeAll = False
End Enum

Enum icdInvoiceType
    AllInvoices = 1004
    MembershipOnly = 1
    MembershipAddonOnly = 2
    ServicesOnly = 3
End Enum
