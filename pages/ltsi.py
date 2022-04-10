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
    if st.button("Generate File"):
        if upload_ltsi is None:
            st.error("ERROR: Please upload File")

        def download_file_ltsi(ltsi):
            import io
            # Writing df to Excel Sheet
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                ltsi.to_excel(writer, sheet_name="valid in ltsi", index=False)

                writer.save()

                st.write("Download Completed File:")
                st.download_button(
                    label="Download Excel worksheets",
                    data=buffer,
                    file_name="LTSI_tool_.xlsx",
                    mime="application/vnd.ms-excel"
                )

        if upload_ltsi is not None:
            upload = pd.read_excel(upload_ltsi, sheet_name=0, engine="openpyxl")
            valid = upload[["salesOrderNum"]]
            valid["Valid in LTSI Tool"] = "TRUE"

            valid["salesOrderNum"] = valid["salesOrderNum"].astype(str)
            import re
            valid["salesOrderNum"] = [re.sub(r"[a-zA-Z]", "", x) for x in valid["salesOrderNum"]]
            valid = valid[valid["salesOrderNum"] != '']
            download_file_ltsi(valid)