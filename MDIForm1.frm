VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10635
   ClientLeft      =   2445
   ClientTop       =   1155
   ClientWidth     =   18705
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10815
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   10755
      ScaleWidth      =   18645
      TabIndex        =   0
      Top             =   0
      Width           =   18705
   End
   Begin VB.Menu EMPLOYEEMANAGEMENT 
      Caption         =   "EMPLOYEE  MANAGEMENT"
      Begin VB.Menu ADDEMPLOYEE 
         Caption         =   "ADD EMPLOYEE"
      End
      Begin VB.Menu UPDATEEMPLOYEE 
         Caption         =   "UPDATE EMPLOYEE"
      End
      Begin VB.Menu DELETEEMPLOYEE 
         Caption         =   "DELETE EMPLOYEE"
      End
      Begin VB.Menu SEARCHEMPLOYEE 
         Caption         =   "SEARCH EMPLOYEE"
      End
      Begin VB.Menu VIEWEMPLOYEE 
         Caption         =   "VIEW EMPLOYEE"
      End
   End
   Begin VB.Menu SALARYMANAGEMENT 
      Caption         =   "SALARY MANAGEMENT"
      Begin VB.Menu SALARYDETAILS 
         Caption         =   "SALARY DETAILS"
      End
      Begin VB.Menu VIEWSALARY 
         Caption         =   "VIEW SALARY"
      End
   End
   Begin VB.Menu PACKAGES 
      Caption         =   "PACKAGES"
      Begin VB.Menu ADDPACKAGE 
         Caption         =   "ADD PACKAGE"
      End
      Begin VB.Menu UPDATEPACKAGE 
         Caption         =   "UPDATE PACKAGE"
      End
      Begin VB.Menu DELETEPACKAGES 
         Caption         =   "DELETE PACKAGES"
      End
      Begin VB.Menu VIEWPACKAGE 
         Caption         =   "VIEW PACKAGE"
      End
   End
   Begin VB.Menu VEHICLES 
      Caption         =   "VEHICLES"
      Begin VB.Menu ADDVEHICLES 
         Caption         =   "ADD VEHICLES"
      End
      Begin VB.Menu SEARCHVEHICLES 
         Caption         =   "SEARCH VEHICLES"
      End
      Begin VB.Menu VIEWVEHICLES 
         Caption         =   "VIEW VEHICLES"
      End
   End
   Begin VB.Menu CUSTOMER 
      Caption         =   "CUSTOMER"
      Begin VB.Menu ADDCUSTOMER 
         Caption         =   "ADD CUSTOMER"
      End
      Begin VB.Menu UPDATECUSTOMER 
         Caption         =   "UPDATE CUSTOMER"
      End
      Begin VB.Menu DELETECUSTOMER 
         Caption         =   "DELETE  CUSTOMER"
      End
      Begin VB.Menu VIEWCUSTOMER 
         Caption         =   "VIEW CUSTOMER"
      End
   End
   Begin VB.Menu BOOKINGS 
      Caption         =   "BOOKINGS"
      Begin VB.Menu NEWBOOKING 
         Caption         =   "NEW BOOKING"
      End
      Begin VB.Menu DELETEBOOKING 
         Caption         =   "DELETE BOOKING"
      End
      Begin VB.Menu VIEWBOOKINGS 
         Caption         =   "VIEW BOOKINGS"
      End
   End
   Begin VB.Menu REPORT 
      Caption         =   "REPORT"
      Begin VB.Menu EMPLOYEEREPORT 
         Caption         =   "EMPLOYEE REPORT"
      End
      Begin VB.Menu CUSTOMERREPORT 
         Caption         =   "CUSTOMER REPORT"
      End
      Begin VB.Menu BOOKINGREPORT 
         Caption         =   "BOOKING REPORT"
      End
   End
   Begin VB.Menu NOTIFICATIONS 
      Caption         =   "NOTIFICATIONS"
      Begin VB.Menu ADDNOTIFICATION 
         Caption         =   "ADD NOTIFICATION"
      End
      Begin VB.Menu UPDATENOTIFICATION 
         Caption         =   "UPDATE NOTIFICATION"
      End
      Begin VB.Menu DELETENOTIFICATION 
         Caption         =   "DELETE NOTIFICATION"
      End
      Begin VB.Menu VIEWNOTIFICATION 
         Caption         =   "VIEW NOTIFICATION"
      End
   End
   Begin VB.Menu MANAGEACCOUNTS 
      Caption         =   "MANAGE ACCOUNTS"
      Begin VB.Menu CREATEACCOUNT 
         Caption         =   "CREATE ACCOUNT"
      End
      Begin VB.Menu DELETEACCOUNT 
         Caption         =   "DELETE ACCOUNT"
      End
      Begin VB.Menu CHANGEPASSWORD 
         Caption         =   "CHANGE PASSWORD"
      End
   End
   Begin VB.Menu LOGOUT 
      Caption         =   "LOG OUT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ADDCUSTOMER_Click()
add_customer.Show
End Sub

Private Sub ADDEMPLOYEE_Click()
add_employee.Show

End Sub

Private Sub ADDNOTIFICATION_Click()
add_notification.Show
End Sub

Private Sub ADDPACKAGE_Click()
add_package.Show
End Sub

Private Sub ADDVEHICLES_Click()
vehicle_registration.Show
End Sub

Private Sub BOOKINGREPORT_Click()
booking_report.Show
End Sub

Private Sub CHANGEPASSWORD_Click()
change_password.Show
End Sub

Private Sub CREATEACCOUNT_Click()
create_account.Show
End Sub

Private Sub CUSTOMERREPORT_Click()
customer_report.Show
End Sub

Private Sub DELETEACCOUNT_Click()
delete_account.Show
End Sub

Private Sub DELETEBOOKING_Click()
delete_booking.Show
End Sub

Private Sub DELETECUSTOMER_Click()
delete_customer.Show
End Sub

Private Sub DELETEEMPLOYEE_Click()
delete_employee.Show
End Sub

Private Sub DELETENOTIFICATION_Click()
delete_nitification.Show
End Sub

Private Sub DELETEPACKAGES_Click()
DELETE_PACKAGE.Show
End Sub


Private Sub EMPLOYEEREPORT_Click()
employee_report.Show
End Sub

Private Sub LOGOUT_Click()
Unload Me
login.Show

End Sub

Private Sub NEWBOOKING_Click()
new_booking.Show
End Sub

Private Sub SALARYDETAILS_Click()
salary_details.Show
End Sub

Private Sub SEARCHEMPLOYEE_Click()
search_employee.Show
End Sub

Private Sub SEARCHVEHICLES_Click()
search_vehicle.Show
End Sub

Private Sub UPDATECUSTOMER_Click()
update_customer.Show
End Sub

Private Sub UPDATEEMPLOYEE_Click()
update_employee.Show
End Sub

Private Sub UPDATENOTIFICATION_Click()
update_notification.Show
End Sub

Private Sub UPDATEPACKAGE_Click()
update_package.Show
End Sub

Private Sub VIEWBOOKINGS_Click()
booking_view.Show
End Sub

Private Sub VIEWCUSTOMER_Click()
customer_view.Show
End Sub

Private Sub VIEWEMPLOYEE_Click()
employee_view.Show
End Sub

Private Sub VIEWNOTIFICATION_Click()
notificatiom_view.Show
End Sub

Private Sub VIEWPACKAGE_Click()
package_view.Show
End Sub

Private Sub VIEWSALARY_Click()
salary_view.Show
End Sub

Private Sub VIEWVEHICLES_Click()
vehicle_view.Show
End Sub
