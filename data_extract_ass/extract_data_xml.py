from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd

def get_vch_no(obj):
    try:
        return obj.find('VOUCHERNUMBER').text
    except:
        return ""

def get_party(obj, trans_type='Parent'):
    try:
        if trans_type == 'Parent':
            return obj.find('PARTYLEDGERNAME').text
        else:
            return obj.find('LEDGERNAME').text
    except:
        return ""

def get_date(obj):
    try:
        datetime_str = obj.find('DATE').text
        datetime_object = datetime.strptime(datetime_str, '%Y%m%d')
        date_str = datetime_object.strftime("%d-%m-%Y")
        return date_str
    except:
        return ""

def get_ref_no(obj, trans_type):
    try:
        if trans_type == 'Child':
            return obj.find('NAME').text
        else:
            return "NA"
    except:
        return ""

def get_ref_type(obj, trans_type):
    try:
        if trans_type == 'Child':
            return obj.find('BILLTYPE').text
        else:
            return "NA"
    except:
        return ""

def get_ref_date(obj, trans_type):
    try:
        if trans_type == 'Child':
            return ""
        else:
            return "NA"
    except:
        return ""

def get_debtor(obj, trans_type='Parent'):
    try:
        if trans_type == 'Parent':
            return obj.find('PARTYLEDGERNAME').text
        else:
            return obj.find('LEDGERNAME').text
    except:
        return ""

def get_ref_amount(obj, trans_type):
    try:
        if trans_type == 'Child':
            return obj.find('AMOUNT').text
        else:
            return "NA"
    except:
        return ""

def get_amount(obj, trans_type):
    try:
        if trans_type == 'Child':
            return "NA"
        else:
            return obj.find('AMOUNT').text
    except:
        return ""

def get_amount_verify(obj):
    try:
        return obj.find('VOUCHERNUMBER').text
    except:
        return ""

def return_empty():
    return {
        'Date': '', 'Transaction Type': '', 'Vch No.': '', 'Ref No': '', 'Ref Type': '', 'Ref Date': '', 'Debtor': '', 'Ref Amount': '', 'Amount': '', 'Particulars': '', 'Vch Type': '', 'Amount Verified': ''
    }

def write_output(date, trans_type, vch_no, ref_no, ref_type, ref_date, 
                 debtor, ref_amount, amount, party_name, vch_type, amnt_vf):
    data = return_empty()
    data["Date"] = date
    data["Transaction Type"] = trans_type
    data["Vch No."] = vch_no
    data["Ref No"] = ref_no
    data["Ref Type"] = ref_type
    data["Ref Date"] = ref_date
    data["Debtor"] = debtor
    data["Ref Amount"] = ref_amount
    data["Amount"] = amount
    data["Particulars"] = party_name
    data["Vch Type"] = vch_type
    data["Amount Verified"] = amnt_vf

    return data

def get_reference_data(ref_obj, ref_obj2, trans_type):
    # NA for parent and other
    ref_no = get_ref_no(ref_obj, trans_type)
    ref_type = get_ref_type(ref_obj, trans_type)
    ref_date  = get_ref_date(ref_obj, trans_type)
    ref_amount = get_ref_amount(ref_obj, trans_type)

    debtor = get_debtor(ref_obj2, trans_type)
    amount = get_amount(ref_obj2, trans_type)

    return ref_no, ref_type, ref_date, ref_amount, debtor, amount

def extract_data(file_path):
    all_output = []
    try:
        fd = open(file_path, 'r')
        data = fd.read()

        soup = BeautifulSoup(data,'xml')
        vouchers = soup.find_all('VOUCHER')

        for voucher in vouchers:
            try:
                all_data = []
                if voucher.find('VOUCHERTYPENAME').text == 'Receipt':
                    date = get_date(voucher)
                    vch_no = get_vch_no(voucher)
                    vch_type = "Receipt"

                    trans_type = "Parent"
                    ref_no, ref_type, ref_date, ref_amount, debtor, parent_amount = get_reference_data(voucher, voucher, trans_type)
                    party_name = debtor
                    parent_amnt_vf = ""
                    
                    # append data
                    parent_out = write_output(date, trans_type, vch_no, ref_no, ref_type, ref_date, 
                            debtor, ref_amount, parent_amount, party_name, vch_type, parent_amnt_vf)

                    child_amount = 0
                    all_ledger_entries = voucher.find_all('ALLLEDGERENTRIES.LIST')
                    for ledger in all_ledger_entries:            
                        # For other Transaction Type
                        if len(ledger.find('BANKALLOCATIONS.LIST')) > 1:
                            trans_type = "Other"
                            ref_no, ref_type, ref_date, ref_amount, debtor, amount = get_reference_data(voucher, ledger, trans_type)
                            party_name = debtor
                            amnt_vf = "NA"

                            # append data
                            out = write_output(date, trans_type, vch_no, ref_no, ref_type, ref_date, 
                                    debtor, ref_amount, amount, party_name, vch_type, amnt_vf)
                            all_data.append(out)
                        # For child Transaction Type
                        elif len(ledger.find('BILLALLOCATIONS.LIST')) > 1:
                            all_childs = ledger.find_all('BILLALLOCATIONS.LIST')
                            child_amount = 0
                            for child in all_childs:
                                trans_type = "Child"
                                ref_no, ref_type, ref_date, ref_amount, debtor, amount = get_reference_data(child, ledger, trans_type)
                                party_name = debtor
                                amnt_vf = "NA"

                                # append data
                                out = write_output(date, trans_type, vch_no, ref_no, ref_type, ref_date, 
                                        debtor, ref_amount, amount, party_name, vch_type, amnt_vf)
                                all_data.append(out)

                                try:
                                    child_amount+=float(ref_amount)
                                except:
                                    child_amount+=0


                    if float(parent_amount) == child_amount:
                        parent_amnt_vf = "Yes"
                    else:
                        parent_amnt_vf = "No"

                    parent_out['Amount Verified'] = parent_amnt_vf
                    all_output.append(parent_out)
                    all_output.extend(all_data)
            except Exception as ee:
                print("Exception while extracting voucher:", ee)

    except Exception as e:
        print("Exception while extracting data:", e)

    df = pd.DataFrame(data=all_output)
    print("DONE")
    df.to_excel('Results.xlsx', index=False)
    return all_output, df
