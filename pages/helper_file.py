
import streamlit as st
import pandas as pd


def app():
    st.write("""

        # LTSI Tool \n

        ### For your second upload please upload your raw download for the day \n \n \n

        ### Contact me if issues arise:
        Slack: @Cameron Looney \n
        email: cameron_j_looney@apple.com""")

    # Button to start the process

    upload_ltsi = st.file_uploader("Upload Raw LTSI Status File", type="xlsx")
    upload_previous_open_orders = st.file_uploader("Upload Yesterdays Open Orders", type="xlsx")
    upload_previous_helper = st.file_uploader("Upload Yesterdays Helper File", type="xlsx")
    if st.button("Generate File"):
        if upload_ltsi is None:
            st.error("ERROR: Please upload File")

        def download_file(ltsi, feedback):
            import io
            # Writing df to Excel Sheet
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                vlookup = pd.read_excel(upload_previous_helper, sheet_name=0, engine="openpyxl")

                dropdown = pd.read_excel(upload_previous_helper, sheet_name=2, engine="openpyxl")

                frames = {'vlookup': vlookup, 'previous feedback': feedback,
                          'dropdown': dropdown, "valid in ltsi": ltsi}

                # now loop thru and put each on a specific sheet
                for sheet, frame in frames.items():  # .use .items for python 3.X
                    frame.to_excel(writer, sheet_name=sheet, index=False)
                formatdict = {'num_format': 'dd/mm/yyyy'}
                workbook = writer.book
                worksheet = writer.sheets['vlookup']
                fmt = workbook.add_format(formatdict)
                worksheet.set_column('C:C', None, fmt)
                # critical last s

                writer.save()

                st.write("Download Completed File:")
                st.download_button(
                    label="Download Excel worksheets",
                    data=buffer,
                    file_name="LTSI_tool_.xlsx",
                    mime="application/vnd.ms-excel"
                )

        if upload_ltsi is not None and upload_previous_open_orders is not None:
            upload = pd.read_excel(upload_ltsi, sheet_name=0, engine="openpyxl")
            open_orders = pd.read_excel(upload_previous_open_orders, sheet_name=0, engine="openpyxl")
            helper = pd.read_excel(upload_previous_helper, sheet_name=0, engine="openpyxl")
            valid = upload[["salesOrderNum"]]
            valid["Valid in LTSI Tool"] = "TRUE"

            valid["salesOrderNum"] = valid["salesOrderNum"].astype(str)
            import re
            valid["salesOrderNum"] = [re.sub(r"[a-zA-Z]", "", x) for x in valid["salesOrderNum"]]
            valid = valid[valid["salesOrderNum"] != '']

            status_col_num = open_orders.columns.get_loc("Status (SS)")
            yesterday = open_orders.iloc[:, [8, status_col_num, 37, 38, 39]]
            download_file(valid, yesterday)