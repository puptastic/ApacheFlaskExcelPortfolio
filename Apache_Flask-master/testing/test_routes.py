from unittest import TestCase
import pytest
from flask import Flask
import json
from app import app


@pytest.fixture
def client():  # fixes issue with 'double' client phenomenon
    with app.test_client() as client:
        yield client


class Testing:
    def test_forward_slash(self, client):
        url = '/'

        response = client.get(url)

        assert response.status_code == 200


    def test_run_a_report(self, client):
        url = '/run_a_report'

        mock_request_headers = {
            'authorization-sha256': 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'
        }

        mock_request_data = {
            'report_id': '12',
            'table_id': 'bnvqs469i'
        }

        response = client.post(url, data=json.dumps(mock_request_data), headers=mock_request_headers)
        assert response.status_code == 200

    def test_run_a_report_b(self, client):
        url = '/run_a_report'

        mock_request_headers = {
            'authorization-sha256': 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'
        }

        mock_request_data = {
            'report_id': '12',
            'table_id': 'bnvqs469i',
            'qb_realm': 'SECRET SECRET SECRET',
            'auth_token': 'SECRET SECRET SECRET'
        }

        response = client.post(url, data=json.dumps(mock_request_data), headers=mock_request_headers)
        assert response.status_code == 200

    def test_process_excel_color(self, client):
        url = '/qb_report_to_color_excel'

        mock_request_headers = {
            'authorization-sha256': 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'
        }

        mock_request_data = {
            'report_id': '12',
            'table_id': 'bnvqs469i',
            'qb_realm': 'SECRET SECRET SECRET',
            'qb_token': 'SECRET SECRET SECRET',
            'file_name': 'testing_skunkworks_from_pycharm',
            'dest_table_id': 'br3imf7tm',
            'dest_rec_id': '1',
            'dest_fid': '6'
        }

        response = client.post(url, data=json.dumps(mock_request_data), headers=mock_request_headers)
        assert response.status_code == 200

    def test_process_excel_color_b(self, client):
        url = '/qb_report_to_color_excel'

        mock_request_headers = {
            'authorization-sha256': 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'
        }

        mock_request_data = {
            'report_id': '12',
            'table_id': 'bnvqs469i',
            'qb_realm': 'SECRET SECRET SECRET',
            'auth_token': 'SECRET SECRET SECRET',
        }

        response = client.post(url, data=json.dumps(mock_request_data), headers=mock_request_headers)
        assert response.status_code == 200

    # For testing on client app
    def test_process_excel_color_c(self, client):
        url = '/qb_report_to_color_excel'

        mock_request_headers = {
            'authorization-sha256': 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'
        }

        mock_request_data = {
            'report_id': '5',
            'table_id': 'bn5w3s2ft',
            'qb_realm': 'SECRET SECRET SECRET',
            'qb_token': 'SECRET SECRET SECRET',
            'file_name': 'testing_from_pycharm',
            'dest_table_id': 'bsbdrswyk',
            'dest_rec_id': '1',
            'dest_fid': '9'
        }

        response = client.post(url, data=json.dumps(mock_request_data), headers=mock_request_headers)
        assert response.status_code == 200

    def test_process_excel_color_d(self, client):
        url = '/qb_report_to_color_excel'

        mock_request_headers = {
            'authorization-sha256': 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'
        }

        mock_request_data = {
            'report_id': '99',
            'table_id': 'bqry48tt3',
            'qb_realm': 'SECRET SECRET SECRET',
            'qb_token': 'SECRET SECRET SECRET',
            'file_name': 'testing_from_pycharm',
            'dest_table_id': 'br5tyfuj5',
            'dest_rec_id': '1',
            'dest_fid': '9'
        }

        response = client.post(url, data=json.dumps(mock_request_data), headers=mock_request_headers)
        assert response.status_code == 200


    def test_skunkworks_table_insert(self, client):
        url = '/qb_insert_skunkworks_record'

        mock_request_headers = {
            'authorization-sha256': 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'
        }

        mock_request_data = {
            'table_id': 'bnvqs469i',
            'qb_realm': 'SECRET SECRET SECRET',
            'auth_token': 'SECRET SECRET SECRET'
        }

        response = client.post(url, data=json.dumps(mock_request_data), headers=mock_request_headers)
        assert response.status_code == 200


    def test_insert_qb_record_x(self, client):
        url = '/insert_qb_record_x'

        mock_request_headers = {
            'authorization-sha256': 'LH_98AH94_HTHASPDTY_9834YTPHAPSDHU'
        }

        mock_request_data = {
            'table_id': 'bnvqs469i',
            'qb_realm': 'SECRET SECRET SECRET',
            'auth_token': 'SECRET SECRET SECRET',
            'field_values': [
                {
                    '8': {
                        'value': 'test_insert_qb_record_x'
                    },
                    '17': {
                        'value': 23
                    }
                }
            ]
        }

        response = client.post(url, data=json.dumps(mock_request_data), headers=mock_request_headers)
        assert response.status_code == 200