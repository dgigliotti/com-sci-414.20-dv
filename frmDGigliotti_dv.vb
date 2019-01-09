' Assignment 6
' Daniela Gigliotti
' COM SCI X 414.20 (Fall 2018)

Public Class frmDesertVistas

    ' Declaration of variables and constants
    Const intMinBase As Integer = 1000    'Minimum Base Rent value that a user can enter
    Const intPrem2Bed As Integer = 150 ' Premium for 2 Bedroom bedroom option - as additional cost
    Const intPremCasita As Integer = 96 ' Premium for Casita bedroom option - as additional cost
    Const intPremLoft As Integer = 2 ' Premium for Loft bedroom option - as multiplier
    Const intPremHalfBath As Integer = 60 ' Premium for 1 1/2 Bathrooms if not standard - as additional cost 
    Const IntPrem2FullBath As Integer = 37 ' Prenium for 2 Full Bathrooms if not standard - as additional cost
    Const intPremGreenbelt As Integer = 6 ' Premium for Greenbelt view - as percentage of RBA
    Const intPremSanJacinto As Integer = 15 ' Premium for San Jacinto Mountains view - as percentage of RBA
    Const intPremGolfCourse As Integer = 33 ' Premium for Golf Course view - as percentage of RBA
    Const intPremLakes As Integer = 33 ' Premium for Lakes view - as percentage of RBA
    Const intPremSantaRosa As Integer = 33 ' Premium for Santa Rosa Mountainse view - as percentage of RBA
    Const intDiscNoLanai As Integer = 42 ' Discount for foregoing Lanai amenity - as subtracted amount
    Const intPremPet As Integer = 25 ' Premium for Pet amenity - as additional amount
    Const intPremGolfCart As Integer = 70 ' Premium for Golf Cart Space amenitye
    Const intPremPool As Integer = 10 ' Premium for Pool Proximity amenity
    Const intPetDeposit As Integer = 200 ' Additional deposit amount for Pet amenity
    Const intGolfCartDeposit As Integer = 50 ' Additional deposit amount for Golf Cart Space amenity
    Const intBadCreditDeposit As Integer = 2 ' Additional deposit amount for Bad Credit option - as multiplier
    Dim dblBase As Integer    ' Base Rent Value
    Dim dblRBA As Double ' Total Rent Before Amenities
    Dim dblTotal As Double ' Total Monthly Rent value after amenities
    Dim dblDeposit As Double ' Total Deposit amount after amenities

    ' When form loads
    Private Sub frmDesertVistas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetupDefaults() ' Set up default options
    End Sub

    'When Tahquitz area option is checked
    Private Sub rdoTahquitz_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTahquitz.CheckedChanged
        EnableOptTahquitz() ' Enable other options available for Tahquitz area
    End Sub

    'When Palm Island area option is checked
    Private Sub rdoPalmIsland_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPalmIsland.CheckedChanged
        EnableOptPalmIsland() ' Enable other options available for Palm Island area option
    End Sub

    'When Moderne area option is checked
    Private Sub rdoModerne_CheckedChanged(sender As Object, e As EventArgs) Handles rdoModerne.CheckedChanged
        EnableOptModerne() ' Enable other options available for Moderne area option
    End Sub

    ' When 1 Bed bedroom option is checked
    Private Sub rdo1Bed_CheckedChanged(sender As Object, e As EventArgs) Handles rdo1Bed.CheckedChanged
        EnableOpt1Bed() ' Enable other options available for 1 Bed bedroom option
    End Sub

    ' When 1 Bed bedroom option is checked
    Private Sub rdo2Bed_CheckedChanged(sender As Object, e As EventArgs) Handles rdo2Bed.CheckedChanged
        EnableOpt2Bed() ' Enable other options available for 2 Bed bedroom option
    End Sub

    ' When Casita bedroom option is checked
    Private Sub rdoCasita_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCasita.CheckedChanged
        EnableOptCasita() ' Enable other options available for Casita bedroom option
    End Sub

    ' When Loft bedroom option is checked
    Private Sub rdoLoft_CheckedChanged(sender As Object, e As EventArgs) Handles rdoLoft.CheckedChanged
        EnableOptLoft() ' Enable other options available for Loft bedroom option
    End Sub

    ' When user changes option in cboView
    Private Sub cboView_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboView.SelectedIndexChanged
        If rdoModerne.Checked = True Then 'If Moderne area option is selected
            DisableViewModerne() ' Validate new View selection is available for Moderne
        End If
    End Sub

    'When Calculate button is clicked
    Private Sub btnCalc_Click(sender As Object, e As EventArgs) Handles btnCalc.Click

        ' Validate that user has entered appropriate data in txtBaseValue
        If Not IsNumeric(txtBaseValue.Text) Then ' If text user has entered as Base Rent is not numeric
            lblTotalValue.Text = "Error" ' Set lblTotalValue to read as error
            lblDepositValue.Text = "Error"  ' Set lblTotalValue to read  as error
            MessageBox.Show("Base rent must be a number", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtBaseValue.Clear() ' Display message box, clear and refocus text in txtBaseValue
            txtBaseValue.Focus()
        ElseIf txtBaseValue.Text < intMinBase Then ' ' If text user has entered as Base Rent exceeds intMinBase
            lblTotalValue.Text = "Error" ' Set lblTotalValue to read as error
            lblDepositValue.Text = "Error"  ' Set lblTotalValue to read as error
            MessageBox.Show("Base rent must be " & intMinBase & " or higher", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtBaseValue.Clear() ' display message box, clear and refocus text in txtBaseValue
            txtBaseValue.Focus()
        Else ' Once Base Rent user has entered is validated, call subroutines to calculate and display values
            CalculateRBA()  ' Calculate total RBA
            CalculateTotal() ' Calculate total rent value
            CalculateDeposit() ' Calculate total deposit amount
            DisplayTotals() ' Display total rent value and Deposit amount to designated labels
        End If
    End Sub

    'When Exit button is clicked, exit program
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    'When Cozy LoveNest menu option is clicked
    Private Sub mnuCozyLovenest_Click(sender As Object, e As EventArgs) Handles mnuCozyLovenest.Click
        rdoPalmIsland.Checked = True ' Set Palm Island area option to checked
        rdoCasita.Checked = True ' Set Casita bedroom option to checked
        cboView.Text = "Lakes" ' Set Lakes view option to checked

        ' Set all Amenities and Tenant-Specific options to unchecked
        chkLanai.Checked = False
        chkPet.Checked = False
        chkPool.Checked = False
        chkGolfCart.Checked = False
        chkBadCredit.Checked = False
    End Sub

    'When Executive Cool Pad menu option is clicked
    Private Sub mnuExecutiveCoolPad_Click(sender As Object, e As EventArgs) Handles mnuExecutiveCoolPad.Click
        rdoModerne.Checked = True ' Set Moderne area option to checked
        cboView.Text = "San Jacinto Mountains" ' Set San Jacinto Mountains view option to checked
        chkLanai.Checked = True ' Set Lanai amenity option to checked
        chkPet.Checked = True ' Set Pet amenity option to checked
        chkPool.Checked = False ' Set Pool Proximity amentiy option to unchecked
        chkGolfCart.Checked = True ' Set Golf Cart Space amenity option to checked
        chkBadCredit.Checked = False ' Set Bad Credit option to unchecked
    End Sub

    ' When Exit menu option is clicked, exit program
    Private Sub mnuExit_Click(sender As Object, e As EventArgs) Handles mnuExit.Click
        Me.Close()
    End Sub

    ' Definition of SetDefaults subroutine
    Private Sub SetupDefaults()
        txtBaseValue.Text = intMinBase ' Set txtBaseValue to intMinBase
        rdoTahquitz.Checked = True ' Set Area option  to rdoTahquitz
        rdo1Bed.Checked = True  ' Set Bedroom option to 1 Bed
        cboView.Text = "Parking" ' Set View option to Parking
        chkLanai.Checked = True ' Set Lanai amenity to checked
        rdoCasita.Enabled = False ' Disable Bedroom option for Casita per Tahquitz options specifications
        rdoLoft.Enabled = False ' Disable Bedroom option for Loft per Tahquitz options specifications
        rdo2Full.Enabled = False ' Disable Bathroom option for 2Full per 1 Bedroom options specifications
    End Sub

    ' Definition of EnableOptTahquitz subroutine
    Private Sub EnableOptTahquitz()
        rdo1Bed.Enabled = True ' Enable 1 Bed bedroom option
        rdo2Bed.Enabled = True ' Enable 2 Bed bedroom option
        rdoCasita.Enabled = False ' Disable Casita bedroom option
        rdoLoft.Enabled = False ' Disable Loft bedroom option
        If rdo2Bed.Checked = True Then ' If 2 Bed bedroom option is already checked
            rdo2Bed.Checked = True ' Keep option as is
        Else
            rdo1Bed.Checked = True ' Else, set 1 Bed bedroom option to checked
        End If
    End Sub

    ' Definition of EnableOptPalmIsland subroutine
    Private Sub EnableOptPalmIsland()
        rdo1Bed.Enabled = True ' Enable 1 Bed bedroom option
        rdo2Bed.Enabled = True ' Enable 2 Bed bedroom option
        rdoCasita.Enabled = True  ' Enable Casita bedroom option
        rdoLoft.Enabled = False  ' Disable Loft bedroom option
        If rdo2Bed.Checked = True Then ' If 2 Bed is already checked
            rdo2Bed.Checked = True ' Keep option as is
        ElseIf rdoCasita.Checked = True Then ' If Casita is already checked
            rdoCasita.Checked = True ' Keep option as is
        Else
            rdo1Bed.Checked = True ' Else, set rdo1Bed to checked
        End If
    End Sub

    ' Definition of EnableOptModerne subroutine
    Private Sub EnableOptModerne()
        rdo1Bed.Enabled = False  ' Disable 1 Bed bedroom option
        rdo2Bed.Enabled = False ' Disable 2 Bed bedroom option
        rdoCasita.Enabled = False ' Disable Casita bedroom option
        rdoLoft.Enabled = True ' Enable Loft bedroom option
        rdoLoft.Checked = True  ' Set Loft to checked
        cboView.Text = "San Jacinto Mountains" ' Set default View to San Jacinto Mountains
        If cboView.Text = "Golf Course" Or cboView.Text = "Lakes" Then ' If cboView is set to Golf Course or Lakes
            MessageBox.Show("View not available for Moderne.", "User Input Error.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cboView.Text = "San Jacinto Mountains" 'Display error message box and set value of cboView to San Jacindo Mountains
        End If
    End Sub

    ' Definition of DisableViewModerne subroutine
    Private Sub DisableViewModerne()
        If cboView.Text = "Golf Course" Or cboView.Text = "Lakes" Then ' If cboView is set to Golf Course or Lakes
            MessageBox.Show("View not available for Moderne", "User Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cboView.Text = "San Jacinto Mountains" 'Display error message box and set value of cboView to San Jacindo Mountains
        End If
    End Sub

    ' Definition of EnableOpt1Bed subroutine
    Private Sub EnableOpt1Bed()
        rdo1Full.Enabled = True ' Enable 1 Full Bath bathroom option
        rdo1Half.Enabled = True ' Enable 1 1/2  Bath bathroom option
        rdo2Full.Enabled = False ' Disable 2 Full Bath bathroom option
        If rdo1Half.Checked = True Then ' If 1 1/2 Bath bathroom is already checked
            rdo1Half.Checked = True ' Keep option as is
        Else
            rdo1Full.Checked = True 'Else set 1 Full Bath bathrpom option to checked
        End If
    End Sub

    ' Definition EnableOpt2Bed subroutine
    Private Sub EnableOpt2Bed()
        rdo1Full.Enabled = False ' Disable 1 Full Bath bathroom option
        rdo1Half.Enabled = True ' Enable 1 1/2 Bath bathroom option
        rdo2Full.Enabled = True  ' Enable 2 Full Bath bathroom option
        If rdo2Full.Checked = True Then ' If 2 Full Bath bathroom is already checked
            rdo2Full.Checked = True ' Keep option as is
        Else
            rdo1Half.Checked = True ' Else set 1 1/2 Bath bathroom optio to checked
        End If
    End Sub

    ' Definition of EnableOptCasita subroutine
    Private Sub EnableOptCasita()
        rdo1Full.Enabled = True  ' Enable 1 Full Bath bathroom option
        rdo1Half.Enabled = False  ' Disable 1 1/2 Bath bathroom option
        rdo2Full.Enabled = False  ' Disable 2 Full Bath bathroom option
        rdo1Full.Checked = True ' Set 1 Full Bath bathroom to checked
    End Sub

    ' Definition of EnableOptLoft subroutine
    Private Sub EnableOptLoft()
        rdo1Full.Enabled = True  ' Enable 1 Full Bath bathroom option
        rdo1Half.Enabled = False  ' Disable 1 1/2 Bath bathroom option
        rdo2Full.Enabled = False  ' Disable 2 Full Bath bathroom option
        rdo1Full.Checked = True ' Set 1 Full Bath bathroom to checked
    End Sub

    'Definition of CalculateRBA subroutine
    Private Sub CalculateRBA()
        dblRBA = txtBaseValue.Text 'Pass Base Rent value to dblRBA once validated

        ' Add Bedroom premium to Total Rent if they apply
        If rdo2Bed.Checked = True Then ' If 2 Bed bedroom option is checked
            dblRBA = dblRBA + intPrem2Bed ' Add 2 Bed bedroom premium to Total Rent
        End If
        If rdoCasita.Checked = True Then ' If Casita bedroom option is checked
            dblRBA = dblRBA + intPremCasita ' Add Casita bedroom premium to Total Rent
        End If
        If rdoLoft.Checked = True Then ' If Loft bedroom option is checked
            dblRBA = dblRBA * intPremLoft ' Add Loft bedroom premium to Total Rent
        End If

        ' Add Bathroom premium to Total Rent is additional bathrooms are selected
        If rdo1Bed.Checked = True And rdo1Half.Checked = True Then ' If 1 Bed and 1 1/2 Bath bathroom options are shecked
            dblRBA = dblRBA + intPremHalfBath ' Add 1 1/2 Bath premium to Total Rent
        End If
        If rdo2Bed.Checked = True And rdo2Full.Checked = True Then ' If 2 Bed and 2 Full Bath bathroom options are checked
            dblRBA = dblRBA + IntPrem2FullBath ' Add 2 Full Bath premium to Total Rent
        End If

        ' Add view premium to Total Rent if they apply
        If cboView.Text = "Greenbelt" Then ' If Greenbelt view option is checked
            dblRBA = (dblRBA + (dblRBA * (intPremGreenbelt * 0.01))) ' Add Greenbelt view premium (percentage) to Total Rent
        End If
        If cboView.Text = "San Jacinto Mountains" Then ' If San Jacinto Mountains view option is checked
            dblRBA = (dblRBA + (dblRBA * (intPremSanJacinto * 0.01))) ' Add San Jacinto Mountains view premium (percentage) to Total Rent
        End If
        If cboView.Text = "Golf Course" Then ' If Golf Course view option is checked
            dblRBA = (dblRBA + (dblRBA * (intPremGolfCourse * 0.01))) ' Add Golf Course view premium (percentage) to Total Rent
        End If
        If cboView.Text = "Lakes" Then ' If Lakes view option is checked
            dblRBA = (dblRBA + (dblRBA * (intPremLakes * 0.01))) ' Add Lakes view premium (percentage) to Total Rent
        End If
        If cboView.Text = "Santa Rosa Mountains" Then ' If Santa Rosa Mountains view option is checked
            dblRBA = (dblRBA + (dblRBA * (intPremSantaRosa * 0.01))) ' Add Santa Rosa Mountains view premium (percentage) to Total Rent 
        End If
    End Sub

    'Definition for CalculateTotal subroutine
    Private Sub CalculateTotal()

        dblTotal = dblRBA ' Set TotalRent equal to RBA

        ' Add premiums and discounts for other amenities to RBA if they apply
        If chkLanai.Checked = False Then ' If Lanai amenity option is NOT checked
            dblTotal = dblTotal - intDiscNoLanai ' Add a No Lanai Discount to Total Rent
        End If
        If chkPet.Checked = True Then ' If Pet tenant option is checked
            dblTotal = dblTotal + intPremPet ' Add Pet premium to Total Rent
        End If
        If chkGolfCart.Checked = True Then ' If Golf Cart Space amenity option is checked
            dblTotal = dblTotal + intPremGolfCart ' Add Gold Cart Space premium to Total Rent
        End If
        If chkPool.Checked = True Then ' If Pool Proximity amenity is checked
            dblTotal = dblTotal + intPremPool ' Add Pool Proximity premium to Total Rent
        End If
    End Sub

    ' Definition for Calculate Deposit subroutine
    Private Sub CalculateDeposit()

        dblDeposit = dblTotal ' Set Deposit equal to Total

        'Add fees to deposit based on amenities chosen
        If chkPet.Checked = True Then ' If Pet tenant option is chedked
            dblDeposit = dblDeposit + intPetDeposit ' Add Pet fee to Deposit
        End If
        If chkGolfCart.Checked = True Then ' If Golf Cart Parking Spac amenitye option is checked
            dblDeposit = dblDeposit + intGolfCartDeposit 'Add Gold Cart Space fee to Deposit
        End If
        If chkBadCredit.Checked = True Then ' If BadCredit option is checked
            dblDeposit = dblDeposit * intBadCreditDeposit ' Add Bad Credit fee (multiplier) to Deposit
        End If
    End Sub

    'Definition of DisplayTotals subroutine
    Private Sub DisplayTotals()
        lblTotalValue.Text = ("$" & Format(Math.Round(dblTotal, 0, MidpointRounding.ToEven), "#,###")) ' Display formatted and rounded (to whole integer) Total Rent Value in lblTotalValue 
        lblDepositValue.Text = ("$" & Format(Math.Round(dblDeposit, 0, MidpointRounding.ToEven), "#,###")) ' Display formatted and rounded (to whole integer) total Deposit amount in lblDepositValue
    End Sub

End Class
