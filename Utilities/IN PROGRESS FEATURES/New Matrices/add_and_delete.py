import os
from collections import defaultdict
#add_new_procedures

def try_int(val):
    if val.strip().isdigit():
        return int(val)
    else:
        return val

def get_all_non_blank_rows(column,rows,rowe):
    return [row for row in range(rows,rowe) if not Cell(row,column).is_empty()]

def get_all_non_blank_values(column,rows,rowe):
    return [str(Cell(row,column).value).strip() for row in range(rows,rowe) if not Cell(row,column).is_empty()]

def get_row_given_procedure(procedure):
    rrange=get_all_non_blank_rows(1,6,last_cell_in_col('A').row-3)
    if 'qam' not in procedure.lower():
        procedure="QAP "+procedure.strip()
    elif procedure.lower().strip()=='qam':
        procedure='QAM 20'
    _,prod_id=procedure.split(" ")
    if "-" in prod_id:
        prod_id=map(try_int,prod_id.split("-"))
    else:
        prod_id=[int(prod_id),None]
    for i,r in enumerate(rrange):
        reference=str(Cell(r,1).value).strip()
        if reference.lower().strip()=='qam':
            reference='QAM 20'
        if reference==procedure:
            return True
        print reference
        _,ref_id=reference.split(" ")
        if "-" in ref_id:
            ref_id=map(try_int,ref_id.split("-"))
        else:
            ref_id=[int(ref_id),None]
        if prod_id[0]<=ref_id[0]:
            if ref_id[1] is None:
                continue
        if i<len(rrange)-1:
            reference2=str(Cell(rrange[i+1],1).value).strip()
            if reference2==procedure:
                return True
            _,ref_id2=reference2.split(" ")
            if "-" in ref_id2:
                ref_id2=map(try_int,ref_id2.split("-"))
            else:
                ref_id2=[ref_id2,None]
            if (ref_id[0]==prod_id[0]==ref_id2[0] and ref_id[1]<=prod_id[1]<ref_id2[1]) or ref_id[0]<=prod_id[0]<ref_id2[0]:
                #print ref_id,prod_id
                if len(ref_id)>len(prod_id) and ref_id[0]==prod_id[0] and ref_id[1]==prod_id[1]:
                    return rrange[i]
                return rrange[i+1]
        else:
            return r+1
        
def delete_all_bad_rows(procedure_list):
    procs=get_all_non_blank_values(1,6,last_cell_in_col('A').row-3)
    rrange=get_all_non_blank_rows(1,6,last_cell_in_col('A').row-3)
    for proc_i,proc in enumerate(procs):
        if proc not in procedure_list:
            if proc_i!=len(procs)-1:
                for i in range(rrange[proc_i],rrange[proc_i+1]:
                    del_row(i)
        

def add_new_row(procedure,p_info_list):
    rrange=get_all_non_blank_rows(1,6,last_cell_in_col('A').row-3)
    proc_row=get_row_given_procedure(procedure)
    if 'qam' not in procedure.lower():
        procedure="QAP "+procedure.strip()
    insert_row(proc_row)
    vals=[procedure]+get_procedure_info(p_info_list)
    for c in range(1,8):
        Cell(proc_row,c).value=vals[c-1]

def get_procedure_info(procedure,info_list):
    info_l=info_list[procedure]
    return [info_l[2],"R"+info_l[0],'',info_l[3],info_l[4],info_l[5]]

def pull_procedure_info():
    caw=active_wkbk()
    plistpath="C:/Users/User/SyncedFolder/Training/Training Program Management(Tree Backup)/tdpstack/general/bnpptraining.xlsx"
    plisttail=os.path.split(plistpath)[1]
    info_dict=defaultdict(list)
    open_wkbk(plistpath)
    active_wkbk(plisttail)
    for row in range(2,last_cell_in_col('A')):
        p=str(Cell(row,1).value)
        if 'qam' not in p:
            p="QAP "+p.strip()
        info_dict[p]=[Cell(row,ci) for ci in range(2,8)]
    active_wkbk(caw)
    return info_dict

def parse_procedure_list(pliststring):
    return pliststring.split(",")

##while True:
##    print get_row_given_procedure(raw_input('> '))

def run_add_and_delete():
    new_list=parse_procedure_list(raw_input('> '))
    proc_list=pull_procedure_info()
    delete_all_bad_rows(new_list)
    for procedure in new_list:
        add_new_row(procedure,proc_list)

