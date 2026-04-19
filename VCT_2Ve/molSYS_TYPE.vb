Module molSYS_TYPE
    Structure STR_CONFIG_BAR
        Dim A As Integer
        Dim Fi As Integer
        Dim N As Integer

        Function WGet_to_string_Fi_A() As String
            Return "D" & Fi & "@" & A
        End Function

        Function WGet_to_string_N_A() As String
            Return N & "D" & Fi
        End Function

        Function WGet_TextNotes_Fi_A() As String
            Return Return_Bar_Notes_Fi_A(Fi, A)
        End Function

        Function WGet_TextNotes_N_Fi() As String
            Return Return_Bar_Notes_N_Fi(N, Fi)
        End Function

        Sub WSet_from_string_Fi_A(ByVal xText As String)
            Dim xText1() As String = xText.Split("@")
            Dim xText2() As String = xText1(0).Split("D")
            A = xText1(1)
            Fi = xText2(1)
        End Sub

        Sub WSet_from_string_N_Fi(ByVal xText As String)
            Dim xText1() As String = xText.Split("D")
            N = xText1(0)
            Fi = xText1(1)
        End Sub

    End Structure
End Module
