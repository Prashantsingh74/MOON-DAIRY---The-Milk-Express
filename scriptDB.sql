create table stock
(
pr_id varchar2(20) constraint stock_fk_productentry references product_entry(pr_id),
pr_name varchar2(20) not null,
mfd date not null,
exp date not null,
balance number(9) default 0
);

create table login
(
user_id varchar2(20) constraint loginpk primary key,
name varchar2(20) not null,
password varchar2(20) not null,
ph_no number(10) not null
);


create table product_entry
(
pr_id varchar2(20) constraint product_entrypk primary key,
pr_name varchar2(20) not null,
weight number(9,2) not null,
unit  varchar2(10) not null,
mrp number(9,2) not null
);


Create table dairy_entry
(
dairy_id varchar2(20) constraint dairy_entrypk primary key,
dairy_nm varchar(20) not null,
phone_no number(10) unique not null,
lic_no  varchar2(25) unique not null,
Address varchar2(50) not null
);


Create table dairy_dues
(
dairy_id varchar2(20) constraint dairydues_fk_dairyentry references dairy_entry(dairy_id),
dues number(9,2) default 0
);


create table dairy_pr
(
 SERIAL_NO NUMBER(6) constraints dairy_prpk primary key,
 Dairy_ID VARCHAR2(20) constraint dairypr_fk_dairyentry references dairy_entry(dairy_id),
 PR_ID VARCHAR2(20)  constraint dairypr_fk_productentry references product_entry(pr_id),
 RATE NUMBER(9,2) not null
);


create table morning_dairysale
(
bill_no number(9) constraint morning_dairysalepk primary key,
bill_date date not null,
centre_id number(20) constraint mdairysale_fk_collectioncentre references collection_centre(centre_id),
dairy_id varchar2(20) constraint mdairysale_fk_dairyentry references dairy_entry(dairy_id),
pr_id varchar2(20) constraint mdairysale_fk_productentry references product_entry(pr_id),
qty number(9,2) not null,
unit varchar2(10) not null,
fat number(7,2) not null,
rate number(7,2) not null,
total_amt number(9,2) not null
);



create table evening_dairysale
(
bill_no number(9) constraint evening_dairysalepk primary key,
bill_date date not null,
centre_id number(20) constraint edairysale_fk_collectioncentre references collection_centre(centre_id),
d_id varchar2(20) constraint edairysale_fk_dairyentry references dairy_entry(dairy_id),
pr_id varchar2(20) constraint edairysale_fk_productentry references product_entry(pr_id),
qty number(9,2) not null,
unit varchar2(10) not null,
fat number(7,2) not null,
rate number(7,2) not null,
total_amt number(9,2) not null
);



create table dairy_payment
(
bill_no number(9) constraint dairy_paymentpk primary key,
Start_Date date not null,
end_date date not null,
dairy_id varchar2(20) constraint dairypayment_fk_dairyentry references dairy_entry(dairy_id),
Morning_qty number(9,2) not null,
evening_qty number(9,2) not null,
Total_Qty number(6,2) not null,
total_amt number(9,2) not null,
total_dues number(9,2) not null,
payment number(9,2) not null,
Dues_left Number(9,2)not null
);


create table collection_centre
(
Centre_id number(20) constraint collection_centrepk primary key,
reg_date date not null,
centre_name varchar2(20) not null,
ph_no number(10) unique not null,
Address varchar2(30) not null
);

create table centre_dues
(
Centre_id number(20) constraint centredues_fk_collectioncentre references collection_centre(centre_id),
Total_dues number(9,2) default 0
);


create table fatlist
(
fat number(5,2) not null,
rate number(7,2) not null
);



CREATE TABLE purchaseorder_details
(
order_no number(5) constraint purchaseorder_pk primary key,
centre_id number(20) constraint porderdetails_fk_collncentre references collection_centre(centre_id),
order_date date not null,
delivery_date date not null,
dairy_id varchar2(20) constraint porderdetails_fk_dairyentry references dairy_entry(dairy_id),
tot_amt number(9,2) not null,
advance Number(9,2) not null,
dues number(9,2) default 0
);

create table PurchaseOrder_pr
(
order_no number(5) constraint porderpr_fk_porderdetails references purchaseorder_details(order_no),
pr_id varchar2(20) constraint porderpr_fk_productentry references product_entry(pr_id),
qty number(5) not null,
Rate number(9,2) not null,
amount number(9,2) not null
);


create table purchasebill_details
(
Bill_no number(6) constraint purchasebill_pk primary key,
Order_no number(6) constraint purchasebill_fk_porderdetails references purchaseorder_details(order_no),
Purc_date date not null,
dairy_id varchar2(20) constraint purchasebill_fk_dairyentry references dairy_entry(dairy_id),
Net_amt number(9,2) not null,
Discount number(6,2) not null,
Pay_amt number(9,2) not null,
Paid number(9,2) not null,
Dues number(9,2) not null
);

create table purchasebill_pr
(
bill_no number(6) constraint pbillpr_fk_pbilldetails references purchasebill_details(bill_no),
pr_id varchar2(20) constraint pbillpr_fk_productentry references product_entry(pr_id),
Qty number(6) not null,
mfd date not null,
exp date not null,
rate number(9,2) not null,
Amount number(9,2) not null
);


create table purchase_return
(
return_no number(9) constraint purchase_returnpk primary key,
return_date date not null,
centre_id number(20) constraint preturn_fk_collectioncentre references collection_centre(centre_id),
bill_no  number(9) constraint purchasereturn_fk_pbilldetails references purchasebill_details(bill_no),
reason varchar2(50) not null,
bill_amt number(9,2) not null,
paid number(9,2) not null,
dues number(9,2) not null
);

create table purchasereturn_pr
(
return_no number(9) constraint preturnpr_fk_purchasereturn references purchase_return(return_no),
pr_id varchar2(20) constraint preturnpr_fk_productentry references product_entry(pr_id),
Qty number(6) not null,
rate number(9,2) not null,
total number(9,2) not null
);




create table customer_entry
(
cust_id varchar2(6) constraint cust_entrypk primary key,
name varchar2(20) not null,
gender varchar2(8) not null,
address varchar2(50) not null,
ph_no number(10) unique not null
);


Create table customer_dues
(
cust_id varchar2(6) constraint customerdues_fk_customerentry references customer_entry(cust_id),
cust_name varchar2(20) not null,
dues number(9,2) not null
);


create table Farmer_Entry
(
F_ID varchar2(6) constraints Farmer_entrypk primary key,
F_name varchar2(20) not null,
Father_name Varchar2(20) not null ,
ph_no number(10) not null,
Aadhar number(12) unique not null,
address varchar2(50) not null,
PR_ID VARCHAR2(20) constraint farmer_entry_fk_productentry references product_entry
);

create table Farmerbank_details  
(
F_id varchar2(6) constraint Fbank_details_fk_Farmer_entry references farmer_entry,
bank_name varchar2(20) not null,
Acc_holderName varchar2(20) not null,
Acc_no Number(16) not null,
IFSC varchar2(10) not null,
Branch_name varchar2(20) not null 
);

Create Table Farmer_dues
(
F_id varchar2(6) constraint Farmer_dues_fk_Farmer_entry references farmer_entry,
Dues number(9,2) not null)
);

create table Farmer_payment
(
Start_Date date not null,
end_date date not null,
F_id varchar2(6) constraint Farmer_fk_payment references farmer_entry(f_id),
Morning_qty number(9,2) not null,
evening_qty number(9,2) not null,
Total_Qty number(6,2) not null,
total_amt number(9,2) not null,
Previous_due number(9,2) not null,
paid number(9,2) not null,
Dues Number(9,2) not null
);


create table Morning_coll
(
Time varchar2(20) not null,
curr_date date not null,
F_id varchar2(6) constraint morning_fk_collection references farmer_entry,
Qty number(6,2) not null,
Fat number(5,2) not null,
rate number(7,2) not null,
amount number(9,2) not null
);


create table Evening_coll
(
Time varchar2(20) not null,
curr_date date not null,
F_id varchar2(6) constraint evening_fk_collection references farmer_entry,
Qty number(6,2) not null,
Fat number(5,2) not null,
rate number(7,2) not null,
amount number(9,2) not null
);


create table salebill_details
(
bill_no number(9) constraint salebillpk primary key,
bill_date date not null,
order_no number(6) constraint salebill_fk_porderdetails references saleorder_details(order_no),
cust_id varchar2(15) constraint salebill_fk_customerentry references customer_entry(cust_id),
tot_amt number(9,2) not null,
discount number(9,2) not null,
bill_amt number(9,2) not null,
paid number(9,2) not null,
dues number(9,2) not null
);


create table salebill_pr
(
bill_no number(9) constraint salebill_fk_salebillproduct references salebill_details(bill_no),
PR_ID VARCHAR2(20) constraint salebill_fk_productentry references product_entry(pr_id),
qty number(9,2) not null,
rate number(9,2) not null,
amount number(9,2) not null
);

CREATE TABLE SALEORDER_details
(
ORDER_NO number(6) constraint saleorderpk primary key,
order_date date not null,
delivery_date date not null,
cust_id varchar2(6) constraint saleorder_fk_customerentry references customer_entry(cust_id),
tot_amt number(9,2) not null,
Paid Number(9,2) not null,
dues number(9,2) not null
);



create table saleorder_pr
(
order_no number(6) constraint sorderpr_fk_sorderdetails references saleorder_details(order_no),
PR_ID VARCHAR2(20) constraint saleorder_fk_productentry references product_entry(pr_id),
qty number(5) not null,
Rate number(9,2) not null,
amount number(9,2) not null
);


create table sales_return
(
return_no number(9) constraint salereturnpk primary key,
return_date date not null,
cust_id varchar2(20) constraint salereturn_fk_customer references customer_entry(cust_id),
bill_no number(9) constraint salereturn_fk_salebill references salebill_details(bill_no),
reason varchar2(50) not null,
net_amt number(9,2) not null,
paid number(9,2) not null,
dues number(9,2) not null
);

create table salesreturn_pr
(
return_no number(9) constraint salereturnpr_fk_salesretun references sales_return(return_no),
pr_id varchar2(10) constraint salereturn_fk_productentry references product_entry(pr_id),
qty number(9,2) not null,
rate number(9,2) not null,
total number(9,2) not null
);