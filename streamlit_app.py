import streamlit as st
import writein

st.title("BSMI 用印文件")

def time_sep(time):
    return str(time).split("-")

with st.form("company_form"):  # 表單名稱可自訂
    # 公司資料
    co_name = st.text_input("公司名稱")
    co_addr = st.text_input("公司地址")
    co_tel = st.text_input("公司電話")
    co_id = st.text_input("公司統編")
    co_represent = st.text_input("公司負責人")

    # 產品資料
    product_name = st.text_input("產品名稱")
    main = st.text_input("主型號")
    series = st.text_input("系列型號")

    # 報告資訊
    report_no = st.text_input("報告編號")
    application_no = st.text_input("受理編號")
    date = st.date_input("簽署日期")  
    review_date = st.date_input("預審日期")

    # 其他欄位
    r_str = st.text_input("R 字軌 (申請者識別號碼)")
    
    test_standard = st.selectbox(
        '測試標準',
        ("CNS 14336-1 資訊技術設備安全通則99年版",
         "CNS 15598-1 影音、資訊及通訊技術設備 -第 1 部：安全要求109年6月30日版",
         "CNS 15425-1(104年版) 電動機車充電系統安全一般規範"),
    )
    unit1 = st.text_input("單元一")
    unit2 = st.text_input("單元二")
    unit3 = st.text_input("單元三")
    unit4 = st.text_input("單元四")


    # 提交按鈕（必須放在 form 裡面）
    submitted = st.form_submit_button("提交")

if submitted:
    Y, M, D = time_sep(date)
    information = {'co_name':co_name, 'co_addr':co_addr, 'co_tel':co_tel, 'co_id':co_id, 'co_represent':co_represent,
                   'product_name':product_name, 'main':main, 'series':series,
                   'report_no':report_no, 'application_no':application_no, 'Y': Y, 'M': M, 'D': D, 'review_date':review_date,
                   'r_str':r_str, 'test_standard':test_standard, 'unit1': unit1, 'unit2': unit2, 'unit3': unit3, 'unit4': unit4}

    # information = {'co_name':'name', 'co_addr':'addr', 'co_tel':'tel', 'co_id':'id', 'co_represent':'represent',
    #             'product_name':'product_name', 'main':'main', 'series':'series',
    #             'report_no':'report_no', 'application_no':'application_no', 'Y': Y, 'M': M, 'D': D, 'review_date':'review_date',
    #             'r_str':'r_str', 'test_standard':'test_standard'}

    zip_buffer = writein.run_BSMI_doc(information)

    st.download_button(
        label="下載壓縮檔",
        data=zip_buffer,
        file_name="results.zip",

        mime="application/zip")








