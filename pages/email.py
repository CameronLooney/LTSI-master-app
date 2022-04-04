import pandas as pd
import streamlit as st
import smtplib

def app():
    #st.set_page_config(page_title='LTSI Emails')
    country_list = ['Portugal','Hungary','UK','Ukraine','Switzerland','CIS','Czech Republic','Russia','Turkey','KSA',
                    'Netherlands','Germany','South Africa','France','Poland','Middle East','Greece / Cyprus','Baltics','Bulgaria / Romania',
                    'Spain','Italy','Sweden','Iceland','Norway','Finland','Denmark','Austria','Israel','Ireland','India','UAE','Belgium / Luxembourg',
                    'Nigeria']
    st.write("""
    
    # LTSI Emails
    
    ### Contact me if issues arise:
    Slack: @Cameron Looney \n
    email: cameron_j_looney@apple.com""")
    col1, col2= st.columns(2)
    with col1:
        st.header('Select email type:')
        option_1 = st.checkbox('All Under Review (both blocked and Reach out)')
        option_2 = st.checkbox('Blocked')
        option_3 = st.checkbox('Reach out to Sales')
    with col2:
        st.header('Countries to Email:')

        options = st.multiselect(
            'Countries selected will be emailed',
            ['All','Portugal',
    'Hungary',
    'UK',
    'Ukraine',
    'Switzerland',
    'CIS',
    'Czech Republic',
    'Russia',
    'Turkey',
    'KSA',
    'Netherlands',
    'Germany',
    'South Africa',
    'France',
    'Poland',
    'Middle East',
    'Greece / Cyprus',
    'Baltics',
    'Bulgaria / Romania',
    'Spain',
    'Italy',
    'Sweden',
    'Iceland',
    'Norway',
    'Finland',
    'Denmark',
    'Austria',
    'Israel',
    'Ireland',
    'India',
    'UAE',
    'Belgium / Luxembourg',
    'Nigeria'],
    ['All'])


    excel_file = st.file_uploader("Upload Excel File", type="xlsx")
    csam_db = st.file_uploader("Upload C-SAM Email List File", type="xlsx")
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(st.secrets["your_email"], st.secrets["your_password"])
    if st.button('Send Emails'):
        if excel_file is not None and csam_db is not None:
            if "All" in options:
                countries_to_email = country_list
            else:
                countries_to_email = list(options)

            try:
                server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
                server.ehlo()
                server.login(st.secrets["your_email"], st.secrets["your_password"])
            except:
                st.error("Could not connect to Server")

            excel_file = pd.read_excel(excel_file, sheet_name=0, engine="openpyxl")
            csam_db = pd.read_excel(csam_db, sheet_name=0, engine="openpyxl")

            feedback_check = [col for col in excel_file.columns if 'Comments' in col]

            unique_countries = list(excel_file["country"].unique())
            df_review = excel_file[excel_file["Status (SS)"] == "Under Review by C-SAM"]
            #df_blocked = excel_file[excel_file["Status (SS)"] == "Blocked"]
            excel_file['del_blk']=excel_file['del_blk'].astype(str)
            excel_file['sch_line_blocked_for_delv'] = excel_file['sch_line_blocked_for_delv'].astype(str)
            ro_codes = ['94','94.0','95','95.0']
            st.write(excel_file['sch_line_blocked_for_delv'].unique())
            df_blocked = excel_file.loc[excel_file['del_blk'].isin(ro_codes) | excel_file['sch_line_blocked_for_delv'].astype(str).isin(ro_codes)]
            st.write(df_blocked)
            # MIGHT NEED TO CHANGE TO JUST SALES ORDER NUMBER ASK JOSEPH WHICH ONE HE SENDS
            import time
            if option_1 and option_2:
                st.error("ERROR: Please choose 1 option")
                time.sleep(2)
                st.experimental_rerun()
            if option_1 and option_3:
                st.error("ERROR: Please choose 1 option")
                time.sleep(2)
                st.experimental_rerun()

            if option_2 and option_3:
                st.error("ERROR: Please choose 1 option")
                time.sleep(2)
                st.experimental_rerun()
            if option_2 and option_3 and option_1:
                st.error("ERROR: Please choose 1 option")
                time.sleep(2)
                st.experimental_rerun()

            if option_1:
                order_num_per_country_under_review = df_review.groupby('country')['Sales Order and Line Item'].agg(
                    set).to_dict()

                countries_to_email_under_review = {k: v for k, v in order_num_per_country_under_review.items() if
                                                   len(v) > 0}

                for key in order_num_per_country_under_review:
                    order_num_per_country_under_review[key] = str(order_num_per_country_under_review[key])
                    order_num_per_country_under_review[key] = str(order_num_per_country_under_review[key]).replace('{','').replace('}','')
                #print(order_num_per_country_under_review)

                # converting dictionary to a dataframe

                # THIS REMOVES COUNTRIES FROM THE DICTIONARY IF THEYARE NOT IN THE LIST OF SELECTED COUNTRIES
                order_num_per_country_under_review= {k: order_num_per_country_under_review[k] for k in countries_to_email if k in order_num_per_country_under_review}
                print(order_num_per_country_under_review)

                dict_dataframe = pd.DataFrame.from_dict(order_num_per_country_under_review, orient='index')
                dict_dataframe.reset_index(inplace=True)
                #print(dict_dataframe)
                dict_dataframe = dict_dataframe.rename(columns={'index': 'country', 0: "codes"})
                merged_under_review = dict_dataframe.merge(csam_db, how='left')
                for index, row in merged_under_review.iterrows():
                    name = row["country"]
                    email = row["email"]
                    subject = "LTSI Orders Under Review"
                    message = (row["codes"])
                    full_email = ("From: {0} <{1}>\n"
                                  "To: {2} <{3}>\n"
                                  "Subject: {4}\n"
                                  "Hello could you please confirm if you want these following orders as they are currently under review.\n"
                                  "{5}\n\n"
                                  "Best Regards,\n"
                                  "Reseller Operations Team\n\n".format(st.secrets["your_name"], st.secrets["your_email"],
                                                                        name, email, subject,
                                                                        message))
                    try:
                        #server.sendmail(st.secrets["your_email"], [email], full_email)
                        st.write('Email to {} successfully sent!\n\n'.format(email))
                    except Exception as e:
                        st.error('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

            if option_2:
                order_num_per_country_blocked = df_blocked.groupby('country')['Sales Order and Line Item'].agg(
                    set).to_dict()

                countries_to_email_blocked = {k: v for k, v in order_num_per_country_blocked.items() if
                                              len(v) > 0}

                for key in order_num_per_country_blocked:
                    order_num_per_country_blocked[key] = str(order_num_per_country_blocked[key])
                    order_num_per_country_blocked[key] = str(order_num_per_country_blocked[key]).replace('{',
                                                                                                         '').replace(
                        '}', '')
                order_num_per_country_blocked = {k: order_num_per_country_blocked[k] for k in countries_to_email
                                                      if k in order_num_per_country_blocked}
                print(order_num_per_country_blocked)
                # converting dictionary to a dataframe
                dict_dataframe_blocked = pd.DataFrame.from_dict(order_num_per_country_blocked, orient='index')
                dict_dataframe_blocked.reset_index(inplace=True)
                dict_dataframe_blocked = dict_dataframe_blocked.rename(columns={'index': 'country', 0: "codes"})
                merged_blocked = dict_dataframe_blocked.merge(csam_db, how='left')
                for index, row in merged_blocked.iterrows():
                    name = row["country"]
                    email = row["email"]
                    subject = "LTSI Orders Blocked"
                    message = (row["codes"])
                    full_email = ("From: {0} <{1}>\n"
                                  "To: {2} <{3}>\n"
                                  "Subject: {4}\n"
                                  "Hello could you please confirm if you want these following orders as they are currently under review.\n"
                                  "{5}\n\n"
                                  "Best Regards,\n"
                                  "Reseller Operations Team\n\n".format(st.secrets["your_name"], st.secrets["your_email"],
                                                                        name, email, subject,
                                                                        message))
                    try:
                        #server.sendmail(st.secrets["your_email"], [email], full_email)
                        st.write('Email to {} successfully sent!\n\n'.format(email))
                    except Exception as e:
                        st.error('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

            if option_3:
                feedback_check = [col for col in excel_file.columns if 'Comments' in col]
                df_reach_out = excel_file[excel_file[feedback_check[0]] == "Reach Out To Sales"]

                if len(feedback_check) == 1:
                    order_num_per_country_reach_out = df_reach_out.groupby('country')['Sales Order and Line Item'].agg(
                        set).to_dict()

                    countries_to_email_reach_out = {k: v for k, v in order_num_per_country_reach_out.items() if
                                                    len(v) > 0}

                    for key in order_num_per_country_reach_out:
                        order_num_per_country_reach_out[key] = str(order_num_per_country_reach_out[key])
                        order_num_per_country_reach_out[key] = str(order_num_per_country_reach_out[key]).replace('{',
                                                                                                                 '').replace(
                            '}',
                            '')
                    order_num_per_country_reach_out = {k: order_num_per_country_reach_out[k] for k in
                                                          countries_to_email if k in order_num_per_country_reach_out}
                    print(order_num_per_country_reach_out)
                    # converting dictionary to a dataframe
                    dict_dataframe_reach_out = pd.DataFrame.from_dict(order_num_per_country_reach_out, orient='index')
                    dict_dataframe_reach_out.reset_index(inplace=True)
                    dict_dataframe_reach_out = dict_dataframe_reach_out.rename(columns={'index': 'country', 0: "codes"})
                    merged_reach_out = dict_dataframe_reach_out.merge(csam_db, how='left')
                    for index, row in merged_reach_out.iterrows():
                        name = row["country"]
                        email = row["email"]
                        subject = "LTSI reaching out for confirmation"
                        message = (row["codes"])
                        full_email = ("From: {0} <{1}>\n"
                                      "To: {2} <{3}>\n"
                                      "Subject: {4}\n"
                                      "Hello could you please confirm if you want these following orders as we were asked to reach out by SDM.\n"
                                      "{5}\n\n"
                                      "Best Regards,\n"
                                      "Reseller Operations Team\n\n".format(st.secrets["your_name"],
                                                                            st.secrets["your_email"], name, email, subject,
                                                                            message))

                        try:
                            #server.sendmail(st.secrets["your_email"], [email], full_email)
                            st.write('Email to {} successfully sent!\n\n'.format(email))
                        except Exception as e:
                            st.error('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

                else:
                    st.error("ERROR: SDM Feedback not availble application restarting")
                    if st.button("Restart"):
                        st.experimental_rerun()



