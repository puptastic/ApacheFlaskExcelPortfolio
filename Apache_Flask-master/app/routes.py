# coding=utf-8
from flask import render_template, flash, redirect, session, url_for, request, g
import json
from app import app
from app import lifters
import requests as RQS


GIBBERISH = 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'


@app.route('/')
def forward_slash():
    return 'Ok', 200


@app.route('/qb_report_to_color_excel', methods=['POST'])
def qb_report_to_color_excel():
    headers = request.headers

    auth_token = headers.get('authorization-sha256')
    if auth_token != GIBBERISH:
        return 'Unauthorized', 401

    data_string = request.get_data()
    data = json.loads(data_string)

    report_id = data.get('report_id')
    table_id = data.get('table_id')
    qb_realm = data.get('qb_realm')
    qb_token = data.get('qb_token')
    file_name = data.get('file_name')

    dest_table_id = data.get('dest_table_id')
    dest_rec_id = data.get('dest_rec_id')
    dest_fid = data.get('dest_fid')

    if report_id and table_id and qb_realm and qb_token:
        if dest_table_id and dest_rec_id and dest_fid:
            process_excel_color = str(lifters.process_excel_color_x(report_id, table_id, qb_realm, qb_token, file_name,
                                                                    dest_table_id, dest_rec_id, dest_fid))
        else:
            process_excel_color = str(lifters.process_excel_color_x(report_id, table_id, qb_realm, qb_token))
    else:
        return 'Bad Request', 400

    return 'Ok', 200


@app.route('/run_a_report', methods=['POST'])
def qb_run_a_report():
    headers = request.headers

    auth_token = headers.get('authorization-sha256')
    if auth_token != GIBBERISH:
        return 'Unauthorized', 401

    data_string = request.get_data()
    data = json.loads(data_string)

    table_id = data.get('table_id')
    report_id = data.get('report_id')
    qb_realm = data.get('qb_realm')
    qb_token = data.get('auth_token')

    if table_id and report_id:
        if qb_realm and auth_token:
            process_result = str(lifters.run_a_report(report_id, table_id, qb_realm, qb_token))
            print(process_result)
        else:
            process_result = str(lifters.run_a_report(report_id, table_id))
            print(process_result)
    else:
        return 'Bad Request', 400

    return 'Ok', 200


@app.route('/qb_insert_skunkworks_record', methods=['POST'])
def qb_insert_skunkworks_record():
    headers = request.headers

    auth_token = headers.get('authorization-sha256')
    if auth_token != GIBBERISH:
        return 'Unauthorized', 401

    data_string = request.get_data()
    data = json.loads(data_string)

    table_id = data.get('table_id')
    qb_realm = data.get('qb_realm')
    qb_token = data.get('auth_token')

    if table_id:
        if qb_realm and auth_token:
            process_result = str(lifters.insert_skunkworks_record(table_id, qb_realm, qb_token))
            print(process_result)
    else:
        return 'Bad Request', 400

    return 'Ok', 200


@app.route('/insert_qb_record_x', methods=['POST'])
def insert_qb_record_x():
    headers = request.headers

    auth_token = headers.get('authorization-sha256')
    if auth_token != GIBBERISH:
        return 'Unauthorized', 401

    data_string = request.get_data()
    data = json.loads(data_string)

    table_id = data.get('table_id')
    qb_realm = data.get('qb_realm')
    qb_token = data.get('auth_token')
    field_values = data.get('field_values')

    if table_id:
        if qb_realm and auth_token:
            process_result = str(lifters.insert_qb_record_x(table_id, qb_realm, qb_token, field_values))
            print(process_result)
    else:
        return 'Bad Request', 400

    return 'Ok', 200


@app.route('/files_to_box', methods=['POST'])
def qb_file_to_box():
    headers = request.headers

    auth_token = headers.get('authorization-sha256')
    if auth_token != GIBBERISH:
        return 'Unauthorized', 401

    data_string = request.get_data()
    data = json.loads(data_string)

    qb_realm = data.get('qb_realm')
    qb_token = data.get('auth_token')
    qb_table_id = data.get('qb_table_id')
    box_url = data.get('box_url')
    box_token = data.get('box_token')

    if qb_realm and qb_token and qb_table_id and box_url and box_token:
        process_result = 5
    else:
        return 'Bad Request', 400

    return 'Ok', 200


@app.route('/get_box_credentials', methods=['POST'])
def get_box_credentials():
    headers = request.headers

    auth_token = headers.get('authorization-sha256')
    if auth_token != GIBBERISH:
        return 'Unauthorized', 401

    data_string = request.get_data()
    data = json.loads(data_string)

    box_client_id = data.get('box_client_id')
    box_client_secret = data.get('box_client_secret')

    process_result = lifters.get_box_credentials(box_client_id, box_client_secret)
    print(process_result)

    return 'Ok', 200
