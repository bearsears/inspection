Attribute VB_Name = "Module1"

Public inspectionStatus As Integer 'this will find if it is satisfactory, failed, or ready to be reinspected.

Sub drawOnMap()
 

    '******************************************
    
    'Display calculation in seconds.
    'Dim StartTime As Double
    'Dim SecondElapsed As Double
    
    'StartTime = Timer
    
    '*****************************************
    'This is the major part of version 2. I want to declare an multi-dimensional string array so
    'that I can start to colour the map.
    
    'For 3, regular expression was implememted to filter out wells by township and range,
    'As well as using autofilter to filter out well status from the start
    'Reduction of map generation time by 30-60% is achieved
    
    
    Dim mapLocation(1 To 36, 1 To 16) As String
  
    'Section 1
    mapLocation(1, 1) = "Y25": mapLocation(1, 2) = "X25": mapLocation(1, 3) = "W25": mapLocation(1, 4) = "V25": mapLocation(1, 5) = "V24": mapLocation(1, 6) = "W24": mapLocation(1, 7) = "X24": mapLocation(1, 8) = "Y24"
    mapLocation(1, 9) = "Y23": mapLocation(1, 10) = "X23": mapLocation(1, 11) = "W23": mapLocation(1, 12) = "V23": mapLocation(1, 13) = "V22": mapLocation(1, 14) = "W22": mapLocation(1, 15) = "X22": mapLocation(1, 16) = "Y22"
    'Section 2
    mapLocation(2, 1) = "U25": mapLocation(2, 2) = "T25": mapLocation(2, 3) = "S25": mapLocation(2, 4) = "R25": mapLocation(2, 5) = "R24": mapLocation(2, 6) = "S24": mapLocation(2, 7) = "T24": mapLocation(2, 8) = "U24"
    mapLocation(2, 9) = "U23": mapLocation(2, 10) = "T23": mapLocation(2, 11) = "S23": mapLocation(2, 12) = "R23": mapLocation(2, 13) = "R22": mapLocation(2, 14) = "S22": mapLocation(2, 15) = "T22": mapLocation(2, 16) = "U22"
    'Section 3
    mapLocation(3, 1) = "Q25": mapLocation(3, 2) = "P25": mapLocation(3, 3) = "O25": mapLocation(3, 4) = "N25": mapLocation(3, 5) = "N24": mapLocation(3, 6) = "O24": mapLocation(3, 7) = "P24": mapLocation(3, 8) = "Q24"
    mapLocation(3, 9) = "Q23": mapLocation(3, 10) = "P23": mapLocation(3, 11) = "O23": mapLocation(3, 12) = "N23": mapLocation(3, 13) = "N22": mapLocation(3, 14) = "O22": mapLocation(3, 15) = "P22": mapLocation(3, 16) = "Q22"
    'Section 4
    mapLocation(4, 1) = "M25": mapLocation(4, 2) = "L25": mapLocation(4, 3) = "K25": mapLocation(4, 4) = "J25": mapLocation(4, 5) = "J24": mapLocation(4, 6) = "K24": mapLocation(4, 7) = "L24": mapLocation(4, 8) = "M24"
    mapLocation(4, 9) = "M23": mapLocation(4, 10) = "L23": mapLocation(4, 11) = "K23": mapLocation(4, 12) = "J23": mapLocation(4, 13) = "J22": mapLocation(4, 14) = "K22": mapLocation(4, 15) = "L22": mapLocation(4, 16) = "M22"
    'Section 5
    mapLocation(5, 1) = "I25": mapLocation(5, 2) = "H25": mapLocation(5, 3) = "G25": mapLocation(5, 4) = "F25": mapLocation(5, 5) = "F24": mapLocation(5, 6) = "G24": mapLocation(5, 7) = "H24": mapLocation(5, 8) = "I24"
    mapLocation(5, 9) = "I23": mapLocation(5, 10) = "H23": mapLocation(5, 11) = "G23": mapLocation(5, 12) = "F23": mapLocation(5, 13) = "F22": mapLocation(5, 14) = "G22": mapLocation(5, 15) = "H22": mapLocation(5, 16) = "I22"
    'Section 6
    mapLocation(6, 1) = "E25": mapLocation(6, 2) = "D25": mapLocation(6, 3) = "C25": mapLocation(6, 4) = "B25": mapLocation(6, 5) = "B24": mapLocation(6, 6) = "C24": mapLocation(6, 7) = "D24": mapLocation(6, 8) = "E24"
    mapLocation(6, 9) = "E23": mapLocation(6, 10) = "D23": mapLocation(6, 11) = "C23": mapLocation(6, 12) = "B23": mapLocation(6, 13) = "B22": mapLocation(6, 14) = "C22": mapLocation(6, 15) = "D22": mapLocation(6, 16) = "E22"
    'Section 7
    mapLocation(7, 1) = "E21": mapLocation(7, 2) = "D21": mapLocation(7, 3) = "C21": mapLocation(7, 4) = "B21": mapLocation(7, 5) = "B20": mapLocation(7, 6) = "C20": mapLocation(7, 7) = "D20": mapLocation(7, 8) = "E20"
    mapLocation(7, 9) = "E19": mapLocation(7, 10) = "D19": mapLocation(7, 11) = "C19": mapLocation(7, 12) = "B19": mapLocation(7, 13) = "B18": mapLocation(7, 14) = "C18": mapLocation(7, 15) = "D18": mapLocation(7, 16) = "E18"
    'Section 8
    mapLocation(8, 1) = "I21": mapLocation(8, 2) = "H21": mapLocation(8, 3) = "G21": mapLocation(8, 4) = "F21": mapLocation(8, 5) = "F20": mapLocation(8, 6) = "G20": mapLocation(8, 7) = "H20": mapLocation(8, 8) = "I20"
    mapLocation(8, 9) = "I19": mapLocation(8, 10) = "H19": mapLocation(8, 11) = "G19": mapLocation(8, 12) = "F19": mapLocation(8, 13) = "F18": mapLocation(8, 14) = "G18": mapLocation(8, 15) = "H18": mapLocation(8, 16) = "I18"
    'Section 9
    mapLocation(9, 1) = "M21": mapLocation(9, 2) = "L21": mapLocation(9, 3) = "K21": mapLocation(9, 4) = "J21": mapLocation(9, 5) = "J20": mapLocation(9, 6) = "K20": mapLocation(9, 7) = "L20": mapLocation(9, 8) = "M20"
    mapLocation(9, 9) = "M19": mapLocation(9, 10) = "L19": mapLocation(9, 11) = "K19": mapLocation(9, 12) = "J19": mapLocation(9, 13) = "J18": mapLocation(9, 14) = "K18": mapLocation(9, 15) = "L18": mapLocation(9, 16) = "M18"
    'Section 10
    mapLocation(10, 1) = "Q21": mapLocation(10, 2) = "P21": mapLocation(10, 3) = "O21": mapLocation(10, 4) = "N21": mapLocation(10, 5) = "N20": mapLocation(10, 6) = "O20": mapLocation(10, 7) = "P20": mapLocation(10, 8) = "Q20"
    mapLocation(10, 9) = "Q19": mapLocation(10, 10) = "P19": mapLocation(10, 11) = "O19": mapLocation(10, 12) = "N19": mapLocation(10, 13) = "N18": mapLocation(10, 14) = "O18": mapLocation(10, 15) = "P18": mapLocation(10, 16) = "Q18"
    'Section 11
    mapLocation(11, 1) = "U21": mapLocation(11, 2) = "T21": mapLocation(11, 3) = "S21": mapLocation(11, 4) = "R21": mapLocation(11, 5) = "R20": mapLocation(11, 6) = "S20": mapLocation(11, 7) = "T20": mapLocation(11, 8) = "U20"
    mapLocation(11, 9) = "U19": mapLocation(11, 10) = "T19": mapLocation(11, 11) = "S19": mapLocation(11, 12) = "R19": mapLocation(11, 13) = "R18": mapLocation(11, 14) = "S18": mapLocation(11, 15) = "T18": mapLocation(11, 16) = "U18"
    'Section 12
    mapLocation(12, 1) = "Y21": mapLocation(12, 2) = "X21": mapLocation(12, 3) = "W21": mapLocation(12, 4) = "V21": mapLocation(12, 5) = "V20": mapLocation(12, 6) = "W20": mapLocation(12, 7) = "X20": mapLocation(12, 8) = "Y20"
    mapLocation(12, 9) = "Y19": mapLocation(12, 10) = "X19": mapLocation(12, 11) = "W19": mapLocation(12, 12) = "V19": mapLocation(12, 13) = "V18": mapLocation(12, 14) = "W18": mapLocation(12, 15) = "X18": mapLocation(12, 16) = "Y18"
    'Section 13
    mapLocation(13, 1) = "Y17": mapLocation(13, 2) = "X17": mapLocation(13, 3) = "W17": mapLocation(13, 4) = "V17": mapLocation(13, 5) = "V16": mapLocation(13, 6) = "W16": mapLocation(13, 7) = "X16": mapLocation(13, 8) = "Y16"
    mapLocation(13, 9) = "Y15": mapLocation(13, 10) = "X15": mapLocation(13, 11) = "W15": mapLocation(13, 12) = "V15": mapLocation(13, 13) = "V14": mapLocation(13, 14) = "W14": mapLocation(13, 15) = "X14": mapLocation(13, 16) = "Y14"
    'Section 14
    mapLocation(14, 1) = "U17": mapLocation(14, 2) = "T17": mapLocation(14, 3) = "S17": mapLocation(14, 4) = "R17": mapLocation(14, 5) = "R16": mapLocation(14, 6) = "S16": mapLocation(14, 7) = "T16": mapLocation(14, 8) = "U16"
    mapLocation(14, 9) = "U15": mapLocation(14, 10) = "T15": mapLocation(14, 11) = "S15": mapLocation(14, 12) = "R15": mapLocation(14, 13) = "R14": mapLocation(14, 14) = "S14": mapLocation(14, 15) = "T14": mapLocation(14, 16) = "U14"
    'Section 15
    mapLocation(15, 1) = "Q17": mapLocation(15, 2) = "P17": mapLocation(15, 3) = "O17": mapLocation(15, 4) = "N17": mapLocation(15, 5) = "N16": mapLocation(15, 6) = "O16": mapLocation(15, 7) = "P16": mapLocation(15, 8) = "Q16"
    mapLocation(15, 9) = "Q15": mapLocation(15, 10) = "P15": mapLocation(15, 11) = "O15": mapLocation(15, 12) = "N15": mapLocation(15, 13) = "N14": mapLocation(15, 14) = "O14": mapLocation(15, 15) = "P14": mapLocation(15, 16) = "Q14"
    'Section 16
    mapLocation(16, 1) = "M17": mapLocation(16, 2) = "L17": mapLocation(16, 3) = "K17": mapLocation(16, 4) = "J17": mapLocation(16, 5) = "J16": mapLocation(16, 6) = "K16": mapLocation(16, 7) = "L16": mapLocation(16, 8) = "M16"
    mapLocation(16, 9) = "M15": mapLocation(16, 10) = "L15": mapLocation(16, 11) = "K15": mapLocation(16, 12) = "J15": mapLocation(16, 13) = "J14": mapLocation(16, 14) = "K14": mapLocation(16, 15) = "L14": mapLocation(16, 16) = "M14"
    'Section 17
    mapLocation(17, 1) = "I17": mapLocation(17, 2) = "H17": mapLocation(17, 3) = "G17": mapLocation(17, 4) = "F17": mapLocation(17, 5) = "F16": mapLocation(17, 6) = "G16": mapLocation(17, 7) = "H16": mapLocation(17, 8) = "I16"
    mapLocation(17, 9) = "I15": mapLocation(17, 10) = "H15": mapLocation(17, 11) = "G15": mapLocation(17, 12) = "F15": mapLocation(17, 13) = "F14": mapLocation(17, 14) = "G14": mapLocation(17, 15) = "H14": mapLocation(17, 16) = "I14"
    'Section 18
    mapLocation(18, 1) = "E17": mapLocation(18, 2) = "D17": mapLocation(18, 3) = "C17": mapLocation(18, 4) = "B17": mapLocation(18, 5) = "B16": mapLocation(18, 6) = "C16": mapLocation(18, 7) = "D16": mapLocation(18, 8) = "E16"
    mapLocation(18, 9) = "E15": mapLocation(18, 10) = "D15": mapLocation(18, 11) = "C15": mapLocation(18, 12) = "B15": mapLocation(18, 13) = "B14": mapLocation(18, 14) = "C14": mapLocation(18, 15) = "D14": mapLocation(18, 16) = "E14"
    'Section 19
    mapLocation(19, 1) = "E13": mapLocation(19, 2) = "D13": mapLocation(19, 3) = "C13": mapLocation(19, 4) = "B13": mapLocation(19, 5) = "B12": mapLocation(19, 6) = "C12": mapLocation(19, 7) = "D12": mapLocation(19, 8) = "E12"
    mapLocation(19, 9) = "E11": mapLocation(19, 10) = "D11": mapLocation(19, 11) = "C11": mapLocation(19, 12) = "B11": mapLocation(19, 13) = "B10": mapLocation(19, 14) = "C10": mapLocation(19, 15) = "D10": mapLocation(19, 16) = "E10"
    'Section 20
    mapLocation(20, 1) = "I13": mapLocation(20, 2) = "H13": mapLocation(20, 3) = "G13": mapLocation(20, 4) = "F13": mapLocation(20, 5) = "F12": mapLocation(20, 6) = "G12": mapLocation(20, 7) = "H12": mapLocation(20, 8) = "I12"
    mapLocation(20, 9) = "I11": mapLocation(20, 10) = "H11": mapLocation(20, 11) = "G11": mapLocation(20, 12) = "F11": mapLocation(20, 13) = "F10": mapLocation(20, 14) = "G10": mapLocation(20, 15) = "H10": mapLocation(20, 16) = "I10"
    'Section 21
    mapLocation(21, 1) = "M13": mapLocation(21, 2) = "L13": mapLocation(21, 3) = "K13": mapLocation(21, 4) = "J13": mapLocation(21, 5) = "J12": mapLocation(21, 6) = "K12": mapLocation(21, 7) = "L12": mapLocation(21, 8) = "M12"
    mapLocation(21, 9) = "M11": mapLocation(21, 10) = "L11": mapLocation(21, 11) = "K11": mapLocation(21, 12) = "J11": mapLocation(21, 13) = "J10": mapLocation(21, 14) = "K10": mapLocation(21, 15) = "L10": mapLocation(21, 16) = "M10"
    'Section 22
    mapLocation(22, 1) = "Q13": mapLocation(22, 2) = "P13": mapLocation(22, 3) = "O13": mapLocation(22, 4) = "N13": mapLocation(22, 5) = "N12": mapLocation(22, 6) = "O12": mapLocation(22, 7) = "P12": mapLocation(22, 8) = "Q12"
    mapLocation(22, 9) = "Q11": mapLocation(22, 10) = "P11": mapLocation(22, 11) = "O11": mapLocation(22, 12) = "N11": mapLocation(22, 13) = "N10": mapLocation(22, 14) = "O10": mapLocation(22, 15) = "P10": mapLocation(22, 16) = "Q10"
    'Section 23
    mapLocation(23, 1) = "U13": mapLocation(23, 2) = "T13": mapLocation(23, 3) = "S13": mapLocation(23, 4) = "R13": mapLocation(23, 5) = "R12": mapLocation(23, 6) = "S12": mapLocation(23, 7) = "T12": mapLocation(23, 8) = "U12"
    mapLocation(23, 9) = "U11": mapLocation(23, 10) = "T11": mapLocation(23, 11) = "S11": mapLocation(23, 12) = "R11": mapLocation(23, 13) = "R10": mapLocation(23, 14) = "S10": mapLocation(23, 15) = "T10": mapLocation(23, 16) = "U10"
    'Section 24
    mapLocation(24, 1) = "Y13": mapLocation(24, 2) = "X13": mapLocation(24, 3) = "W13": mapLocation(24, 4) = "V13": mapLocation(24, 5) = "V12": mapLocation(24, 6) = "W12": mapLocation(24, 7) = "X12": mapLocation(24, 8) = "Y12"
    mapLocation(24, 9) = "Y11": mapLocation(24, 10) = "X11": mapLocation(24, 11) = "W11": mapLocation(24, 12) = "V11": mapLocation(24, 13) = "V10": mapLocation(24, 14) = "W10": mapLocation(24, 15) = "X10": mapLocation(24, 16) = "Y10"
    'Section 25
    mapLocation(25, 1) = "Y9": mapLocation(25, 2) = "X9": mapLocation(25, 3) = "W9": mapLocation(25, 4) = "V9": mapLocation(25, 5) = "V8": mapLocation(25, 6) = "W8": mapLocation(25, 7) = "X8": mapLocation(25, 8) = "Y8"
    mapLocation(25, 9) = "Y7": mapLocation(25, 10) = "X7": mapLocation(25, 11) = "W7": mapLocation(25, 12) = "V7": mapLocation(25, 13) = "V6": mapLocation(25, 14) = "W6": mapLocation(25, 15) = "X6": mapLocation(25, 16) = "Y6"
    'Section 26
    mapLocation(26, 1) = "U9": mapLocation(26, 2) = "T9": mapLocation(26, 3) = "S9": mapLocation(26, 4) = "R9": mapLocation(26, 5) = "R8": mapLocation(26, 6) = "S8": mapLocation(26, 7) = "T8": mapLocation(26, 8) = "U8"
    mapLocation(26, 9) = "U7": mapLocation(26, 10) = "T7": mapLocation(26, 11) = "S7": mapLocation(26, 12) = "R7": mapLocation(26, 13) = "R6": mapLocation(26, 14) = "S6": mapLocation(26, 15) = "T6": mapLocation(26, 16) = "U6"
    'Section 27
    mapLocation(27, 1) = "Q9": mapLocation(27, 2) = "P9": mapLocation(27, 3) = "O9": mapLocation(27, 4) = "N9": mapLocation(27, 5) = "N8": mapLocation(27, 6) = "O8": mapLocation(27, 7) = "P8": mapLocation(27, 8) = "Q8"
    mapLocation(27, 9) = "Q7": mapLocation(27, 10) = "P7": mapLocation(27, 11) = "O7": mapLocation(27, 12) = "N7": mapLocation(27, 13) = "N6": mapLocation(27, 14) = "O6": mapLocation(27, 15) = "P6": mapLocation(27, 16) = "Q6"
    'Section 28
    mapLocation(28, 1) = "M9": mapLocation(28, 2) = "L9": mapLocation(28, 3) = "K9": mapLocation(28, 4) = "J9": mapLocation(28, 5) = "J8": mapLocation(28, 6) = "K8": mapLocation(28, 7) = "L8": mapLocation(28, 8) = "M8"
    mapLocation(28, 9) = "M7": mapLocation(28, 10) = "L7": mapLocation(28, 11) = "K7": mapLocation(28, 12) = "J7": mapLocation(28, 13) = "J6": mapLocation(28, 14) = "K6": mapLocation(28, 15) = "L6": mapLocation(28, 16) = "M6"
    'Section 29
    mapLocation(29, 1) = "I9": mapLocation(29, 2) = "H9": mapLocation(29, 3) = "G9": mapLocation(29, 4) = "F9": mapLocation(29, 5) = "F8": mapLocation(29, 6) = "G8": mapLocation(29, 7) = "H8": mapLocation(29, 8) = "I8"
    mapLocation(29, 9) = "I7": mapLocation(29, 10) = "H7": mapLocation(29, 11) = "G7": mapLocation(29, 12) = "F7": mapLocation(29, 13) = "F6": mapLocation(29, 14) = "G6": mapLocation(29, 15) = "H6": mapLocation(29, 16) = "I6"
    'Section 30
    mapLocation(30, 1) = "E9": mapLocation(30, 2) = "D9": mapLocation(30, 3) = "C9": mapLocation(30, 4) = "B9": mapLocation(30, 5) = "B8": mapLocation(30, 6) = "C8": mapLocation(30, 7) = "D8": mapLocation(30, 8) = "E8"
    mapLocation(30, 9) = "E7": mapLocation(30, 10) = "D7": mapLocation(30, 11) = "C7": mapLocation(30, 12) = "B7": mapLocation(30, 13) = "B6": mapLocation(30, 14) = "C6": mapLocation(30, 15) = "D6": mapLocation(30, 16) = "E6"
    'Section 31
    mapLocation(31, 1) = "E5": mapLocation(31, 2) = "D5": mapLocation(31, 3) = "C5": mapLocation(31, 4) = "B5": mapLocation(31, 5) = "B4": mapLocation(31, 6) = "C4": mapLocation(31, 7) = "D4": mapLocation(31, 8) = "E4"
    mapLocation(31, 9) = "E3": mapLocation(31, 10) = "D3": mapLocation(31, 11) = "C3": mapLocation(31, 12) = "B3": mapLocation(31, 13) = "B2": mapLocation(31, 14) = "C2": mapLocation(31, 15) = "D2": mapLocation(31, 16) = "E2"
    'Section 32
    mapLocation(32, 1) = "I5": mapLocation(32, 2) = "H5": mapLocation(32, 3) = "G5": mapLocation(32, 4) = "F5": mapLocation(32, 5) = "F4": mapLocation(32, 6) = "G4": mapLocation(32, 7) = "H4": mapLocation(32, 8) = "I4"
    mapLocation(32, 9) = "I3": mapLocation(32, 10) = "H3": mapLocation(32, 11) = "G3": mapLocation(32, 12) = "F3": mapLocation(32, 13) = "F2": mapLocation(32, 14) = "G2": mapLocation(32, 15) = "H2": mapLocation(32, 16) = "I2"
    'Section 33
    mapLocation(33, 1) = "M5": mapLocation(33, 2) = "L5": mapLocation(33, 3) = "K5": mapLocation(33, 4) = "J5": mapLocation(33, 5) = "J4": mapLocation(33, 6) = "K4": mapLocation(33, 7) = "L4": mapLocation(33, 8) = "M4"
    mapLocation(33, 9) = "M3": mapLocation(33, 10) = "L3": mapLocation(33, 11) = "K3": mapLocation(33, 12) = "J3": mapLocation(33, 13) = "J2": mapLocation(33, 14) = "K2": mapLocation(33, 15) = "L2": mapLocation(33, 16) = "M2"
    'Section 34
    mapLocation(34, 1) = "Q5": mapLocation(34, 2) = "P5": mapLocation(34, 3) = "O5": mapLocation(34, 4) = "N5": mapLocation(34, 5) = "N4": mapLocation(34, 6) = "O4": mapLocation(34, 7) = "P4": mapLocation(34, 8) = "Q4"
    mapLocation(34, 9) = "Q3": mapLocation(34, 10) = "P3": mapLocation(34, 11) = "O3": mapLocation(34, 12) = "N3": mapLocation(34, 13) = "N2": mapLocation(34, 14) = "O2": mapLocation(34, 15) = "P2": mapLocation(34, 16) = "Q2"
    'Section 35
    mapLocation(35, 1) = "U5": mapLocation(35, 2) = "T5": mapLocation(35, 3) = "S5": mapLocation(35, 4) = "R5": mapLocation(35, 5) = "R4": mapLocation(35, 6) = "S4": mapLocation(35, 7) = "T4": mapLocation(35, 8) = "U4"
    mapLocation(35, 9) = "U3": mapLocation(35, 10) = "T3": mapLocation(35, 11) = "S3": mapLocation(35, 12) = "R3": mapLocation(35, 13) = "R2": mapLocation(35, 14) = "S2": mapLocation(35, 15) = "T2": mapLocation(35, 16) = "U2"
    'Section 36
    mapLocation(36, 1) = "Y5": mapLocation(36, 2) = "X5": mapLocation(36, 3) = "W5": mapLocation(36, 4) = "V5": mapLocation(36, 5) = "V4": mapLocation(36, 6) = "W4": mapLocation(36, 7) = "X4": mapLocation(36, 8) = "Y4"
    mapLocation(36, 9) = "Y3": mapLocation(36, 10) = "X3": mapLocation(36, 11) = "W3": mapLocation(36, 12) = "V3": mapLocation(36, 13) = "V2": mapLocation(36, 14) = "W2": mapLocation(36, 15) = "X2": mapLocation(36, 16) = "Y2"
    
    
    '************************************************************************
        
    'If there is anything on the map delete them all
    Sheets("map").Select
    Sheets("map").range("A1:Y27").ClearContents
    Sheets("map").range("A1:Y27").Interior.ColorIndex = 0
       
    'I need to set up the section township range and meridian
    Dim legalSubDiv As String
    Dim section As String
    Dim township As String
    Dim Ranges As String
    Dim meridian As String
    Dim location As String
    Dim filter As String 'Regular expression
      
    'Get the right address from the menu.
    
    township = Sheets("menu").range("F6")
    Ranges = Sheets("menu").range("F8")
    meridian = Sheets("menu").range("F10")
    
    'This is the addition for 3, filtering out well status.
    Sheets("Wells").Rows("4").AutoFilter
    Sheets("Wells").range("$A$4").AutoFilter Field:=11, Criteria1:=Array _
        ("ACTIVE", "CASED", "COMPLETED", "DEEPENED(DREENTERED)", "DRILLING", _
         "PRESET", "SUSPENDED"), Operator:=xlFilterValues
    'This is the regular expression, regex expression is implemented
    filter = "*" & township & "-" & Ranges & "*"
    Sheets("Wells").range("$A$4").AutoFilter Field:=4, Criteria1:=filter
    
    'The code below will label the map.
    'Sheets("map").Select
    Sheets("map").range("B1") = Ranges & " RGE"
    Sheets("map").range("A2") = township & " TWP"
    Sheets("map").range("B26") = Date
    Sheets("map").range("J26") = township & "-" & Ranges & " " & meridian
    
        
    'The code below will set up variables that I will have to use for rest of the code.
    Dim wellCounter As Integer 'counts number of wells.
    Dim inspectedCounter As Integer 'counts number of wells that have been inspected
    Dim totalWellCounter As Integer 'counts the number of all the wells in the township and range
    Dim totalInspectedCounter As Integer 'counts the number of all the inspected wells in the township and range.
    Dim firstCellAddress As String 'the cell address where the first well using the location have been found
    Dim licence As String 'well licence number
    Dim temp As String 'this is to hold anything temporary.
         
    wellCounter = 0
    inspectedCounter = 0
    totalWellCounter = 0
    totalInspectedCounter = 0
    inspectionStatus = 0 'It is declared as public variable.
       
    With Sheets("Wells").range("D:D")
        For sec = 1 To 36
            'This is to convert the number to the right format of string.
            If sec < 10 Then
                section = "0" & CStr(sec) 'that pesky "0"
            Else
                section = CStr(sec)
            End If
                            
            For lsd = 1 To 16
                'This is to convert the number to the right format of string.
                If lsd < 10 Then
                    legalSubDiv = "0" & CStr(lsd) 'that pesky "0"
                Else
                    legalSubDiv = CStr(lsd)
                End If
                
                'put together all the bits of location information into a full location address
                location = legalSubDiv & "-" & section & "-" & township & "-" & Ranges & meridian
                
                'find the first well, if there is one.
                Set found = .Find(location)
                                
                'The program will enter this logic below if excel did find a first well
                If Not found Is Nothing Then
                    firstCellAddress = found.address

                    'On Gary's request, I am adding the If to filter out the abandoned and planned wells.
                    'it makes sense to filter them out when it comes to for the purpose of inspections.
                    'it is physically hard to find the abandoned wells, or planned wells that have not been drilled.
                    Do
                                          
                        temp = "K" & found.Row

                        
                        temp = found.address
                        

                            'SEARCHING FOR INSPECTED WELL HAS TO START HERE,
                            'MADE A FUNCTION isItInspected THAT WOULD CHECK ON THE SHEET Inspected
                            'WHETHER OR NOT THE WELL IN QUESTION HAS BEEN INSPECTED.
                            'IF IT COMES BACK POSITIVE(ie return 1), IT WILL INCREMENT THE inspectedCounter
                                                   
                            'Get the licence
                        temp = "C" & found.Row
                        licence = Sheets("Wells").range(temp).Value
                            
                        temp = found.address 'reassigning temp because I changed the value in this If nest
                 
                        wellCounter = wellCounter + 1
                        inspectedCounter = inspectedCounter + isItInspected(licence)
                            

                            'Why use find like the one below and not use findnext?
                            'because when using isItInspected in the If nest , the value found changes, it turns to type Nothing.
                            'so i had used temp to store the cell location and start a fresh search after the location.
                        
                        Set found = .Find(location, after:=range(temp))

       
                    Loop While firstCellAddress <> found.address

                    'This If check is for abandoned wells.
                    'If wellCounter > 0 Then temp = placeOnMap(legalSubDiv, section, inspectedCounter, wellCounter, inspectionStatus)
                    If wellCounter > 0 Then
                        Sheets("map").range(mapLocation(sec, lsd)) = inspectedCounter & ":" & wellCounter
                        Sheets("map").range(mapLocation(sec, lsd)).Interior.ColorIndex = inspectionStatus
                    End If
                    
                End If
                
                totalWellCounter = totalWellCounter + wellCounter
                totalInspectedCounter = totalInspectedCounter + inspectedCounter
                wellCounter = 0
                inspectedCounter = 0
                inspectionStatus = 0
            
            Next lsd
              
        Next sec
        
    End With
    
    'Used for debugging
    'MsgBox (totalInspectedCounter)
    'MsgBox (totalWellCounter)
       
    'Below line has been fixed for bug. The If statement had been added to prevent cases
    'where totalWellCounter is 0 in empty maps, and division of zero occurs
    
    If totalWellCounter > 0 Then
        Sheets("map").range("R26") = totalInspectedCounter & " : " & totalWellCounter & " = " & Format((totalInspectedCounter / totalWellCounter), "Percent")
    Else
        Sheets("map").range("R26") = totalInspectedCounter & " : " & totalWellCounter & " = " & Format(0, "Percent")
    End If
    
    Sheets("Wells").Rows("4").AutoFilter
    
    '*******************************************
    'Determine how many seconds code took to run
    'SecondsElapsed = Round(Timer - StartTime, 2)

    'Notify user in seconds
    'MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
    
    

    
          
End Sub

Function isItInspected(licence As String)
    'This is the function that will find out if the well in question has been inspected.
    'I am using the licence as the tool to match.
       
With Sheets("Inspected").range("L:L")
    'Search, but backwards.
    Set c = .Find(licence, after:=Cells(1, 12), searchdirection:=xlPrevious)
    
    Dim temp As String
    Dim status As String
                
    'Enter the instructions below if there is a hit.
    If Not c Is Nothing Then
        
        temp = "R" & c.Row
        status = Sheets("Inspected").range(temp).Value
              
        'Ind for Industry Action Required. 3 is a colour for red.
        If InStr(status, "Ind") Then
            inspectionStatus = 3
        
        'Rea for Ready for Reinspection. 43 is a colour for light green
        ElseIf InStr(status, "Rea") Then
            inspectionStatus = 43
        
        'Inspected and satisfactory. 36 is a colour for whiter yellow.
        Else
            inspectionStatus = 36
     
        End If
        
        isItInspected = 1
        
        Exit Function
        
    End If
    'MsgBox ("not found")

End With

    'no hit has been found.
    isItInspected = 0
    
End Function




