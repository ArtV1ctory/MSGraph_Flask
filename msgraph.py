import socket
import csv
import numpy as np
import adal
import flask
import uuid
import requests
from random import randint

app = flask.Flask(__name__)
app.debug = True
app.secret_key = 'development'

class ExcelClient(object):
  """Class object of ExcelClient"""
  def __init__(self, RESOURCE, TENANT, AUTHORITY_HOST_URL, CLIENT_ID,
               CLIENT_SECRET, API_VERSION, HOST, PORT):
    '''
    Initializes the necessary parameters for interacting with excel online

    :param RESOURCE:            The resource for which access token will be received
    :type:                      str
    :param TENANT:              The tenant name, e.g. contoso.onmicrosoft.com
    :type:                      str
    :param AUTHORITY_HOST_URL:  Default value
    :type:                      str
    :param CLIENT_ID:           The Application ID of your app from your Azure portal
    :type:                      str
    :param CLIENT_SECRET:       The value of the secret key when it was generated on Azure portal
    :type:                      str
    :param API_VERSION:         Web api version
    :type:                      str
    '''
    self.RESOURCE = RESOURCE
    self.TENANT = TENANT
    self.AUTHORITY_HOST_URL = AUTHORITY_HOST_URL
    self.CLIENT_ID = CLIENT_ID
    self.CLIENT_SECRET = CLIENT_SECRET
    self.API_VERSION = API_VERSION
    self.AUTHORITY_URL = AUTHORITY_HOST_URL + '/' + TENANT
    self.TEMPLATE_AUTHZ_URL = ('https://login.microsoftonline.com/{}/oauth2/authorize?' +
                               'response_type=code&client_id={}&redirect_uri={}&' +
                               'state={}&resource={}')
    self.HOST = HOST
    self.PORT = PORT
    self.REDIRECT_URI = 'http://{0}:{1}/getAToken'.format(self.HOST, self.PORT)

  def update_range(self, file_id, sheetname, range, data, format=None,
                   columnHidden=None, formulas=None, formulasLocal=None,
                   formulasR1C1=None, rowHidden=None):
    '''
    Defines the endpoint for sending a patch request to update range.

    :param file_id:        Id excel file
    :type:                 str
    :param sheetname:      The name of the sheet in the excel file for writing data
    :type:                 str
    :param range:          The range of cells for writing data
    :type:                 str
    :param data:           Data required to write to the excel file.
                           The data returned could be of type string, number, or a boolean.
                           Cell that contain an error will return the error string.
    :type:                 list
    :param format:	       Represents Excel's number format code for the given cell
    :type:                 list
    :param columnHidden:   Represents if all columns of the current range are hidden
    :type:                 bool
    :param formulas:     	 Represents the formula in A1-style notation
    :type:                 list
    :param formulasLocal:	 Represents the formula in A1-style notation, in the user's language
                           and number-formatting locale.
                           For example, the English "=SUM(A1, 1.5)" formula
                           would become "=SUMME(A1; 1,5)" in German
    :type:                 list
    :param formulasR1C1:	 Represents the formula in R1C1-style notation
    :type:                 list
    :param rowHidden:	     Represents if all rows of the current range are hidden.
    :type:                 bool
    :return:               The list with the endpoint of the query, the type of the query and the body of the query
    :type:                 list
    '''
    for [k, v] in {'file_id': file_id, 'sheetname': sheetname, 'range': range}.items():
      if type(v) is not str:
        raise TypeError("Invalid {0} type".format(k))
    if data is None:
      raise ValueError("Invalid data")
    size = np.shape(data)
    for [k, v] in {"data": data, "format": format, "formulas": formulas,
                   "formulasLocal": formulasLocal, "formulasR1C1": formulasR1C1}.items():
      if v is not None:
        if type(v) is not list:
          raise TypeError("Invalid {0} type".format(k))
        if np.shape(v) != size:
          raise ValueError("{0} shape and data shape must be the same".format(k))
    for [k, v] in {"columnHidden": columnHidden, "rowHidden": rowHidden}.items():
      if v is not None and type(v) is not bool:
        raise TypeError("Invalid {0} type".format(k))

    request_body = {
      "values": data,
      "numberFormat": format,
      "columnHidden": columnHidden,
      "formulas": formulas,
      "formulasLocal": formulasLocal,
      "formulasR1C1": formulasR1C1,
      "rowHidden": rowHidden
    }
    for [k, v] in request_body.items():
      if v is None:
        del request_body[k]

    return ["/me/drive/items/{0}/workbook/worksheets/{1}/range(address=\'{2}\')".format(file_id, sheetname, range),
            "patch", request_body]

  def get_range(self, file_id, sheetname, range):
    '''
      Defines the endpoint for sending a get request
      to retrieve the properties and relationships of range object.

      :param file_id:   Id excel file
      :type:            str
      :param sheetname: The name of the sheet in the excel file for writing data
      :type:            str
      :param range:     The range of cells for writing data
      :type:            str
      :return:          The list with the endpoint of the query, the type of the query and the body of the query
      :type:            list
    '''
    for [k, v] in {'file_id': file_id, 'sheetname': sheetname, 'range': range}.items():
      if type(v) is not str:
        raise TypeError("Invalid {0} type".format(k))
    return ["/me/drive/items/{0}/workbook/worksheets/{1}/range(address=\'{2}\')".format(file_id, sheetname, range),
            "get", None]

  def insert_empty_cells(self, file_id, sheetname, range, shift="Down"):
    '''
    Defines the endpoint for sending a post request to insert a cell or a range of cells
    into the worksheet in place of this range, and shifts the other cells to make space.
    Request returns a new range object at the now blank space.

    :param file_id:   Id excel file
    :type:            str
    :param sheetname: The name of the sheet in the excel file for writing data
    :type:            str
    :param range:     The range of cells for writing data
    :type:            str
    :param shift:     Specifies which way to shift the cells
                      The possible values are: "Down", "Right"
                      Default value is "Down"
    :type:            str
    :return:          The list with the endpoint of the query, the type of the query and the body of the query
    :type:            list
    '''
    for [k, v] in {'file_id': file_id, 'sheetname': sheetname, 'range': range, 'shift': shift}.items():
      if type(v) is not str:
        raise TypeError("Invalid {0} type".format(k))
    request_body = {
      "shift": shift
    }
    return ["/me/drive/items/{0}/workbook/worksheets/{1}/range(address=\'{2}\')/insert".format(file_id, sheetname, range),
      "post", request_body]

  def clear_range(self, file_id, sheetname, range, applyTo="All"):
    '''
      Defines the endpoint for sending a post request to clear range values, format, fill, border, etc.

      :param file_id:   Id excel file
      :type:            str
      :param sheetname: The name of the sheet in the excel file for writing data
      :type:            str
      :param range:     The range of cells for writing data
      :type:            str
      :param applyTo:   Optional. Determines the type of clear action.
                        The possible values are: "All", "Formats", "Contents"
                        Default value is "All"
      :type:            str
      :return:          The list with the endpoint of the query, the type of the query and the body of the query
      :type:            list

      !!!If successful, this method returns 200 OK response code. It does not return anything in the response body.
      Therefore, you will see a ValueError in the response body.
      '''
    for [k, v] in {'file_id': file_id, 'sheetname': sheetname, 'range': range, 'applyTo': applyTo}.items():
      if type(v) is not str:
        raise TypeError("Invalid {0} type".format(k))
    request_body = {
      "applyTo": applyTo
    }
    return ["/me/drive/items/{0}/workbook/worksheets/{1}/range(address=\'{2}\')/clear".format(file_id, sheetname, range),
      "post", request_body]

  def delete_range(self, file_id, sheetname, range, shift="Up"):
    '''
      Defines the endpoint for sending a post request to delete the cells associated with the range.

      :param file_id:   Id excel file
      :type:            str
      :param sheetname: The name of the sheet in the excel file for writing data
      :type:            str
      :param range:     The range of cells for writing data
      :type:            str
      :param shift:     Specifies which way to shift the cells. The possible values are: "Up", "Left"
                        Default value is "Up"
      :type:            str
      :return:          The list with the endpoint of the query, the type of the query and the body of the query
      :type:            list

      !!!If successful, this method returns 200 OK response code. It does not return anything in the response body.
      Therefore, you will see a ValueError in the response body.
      '''
    for [k, v] in {'file_id': file_id, 'sheetname': sheetname, 'range': range, 'shift': shift}.items():
      if type(v) is not str:
        raise TypeError("Invalid {0} type".format(k))
    request_body = {
      "shift": shift
    }
    return ["/me/drive/items/{0}/workbook/worksheets/{1}/range(address=\'{2}\')/delete".format(file_id, sheetname, range),
      "post", request_body]

  def get_rangeFormat(self, file_id, sheetname, range):
    '''
      Defines the endpoint for sending a get request to retrieve the properties and relationships of range.

      :param file_id:   Id excel file
      :type:            str
      :param sheetname: The name of the sheet in the excel file for writing data
      :type:            str
      :param range:     The range of cells for writing data
      :type:            str
      :return:          The list with the endpoint of the query, the type of the query and the body of the query
      :type:            list
    '''
    for [k, v] in {'file_id': file_id, 'sheetname': sheetname, 'range': range}.items():
      if type(v) is not str:
        raise TypeError("Invalid {0} type".format(k))
    return ["/me/drive/items/{0}/workbook/worksheets/{1}/range(address=\'{2}\')/format".format(file_id, sheetname, range),
      "get", None]

  def get_data(self, path):
    '''
      Converts data from a file in csv format to data in list format
      for insertion into an excel file.

      :param path:  The path to the file in csv format
      :type:        str
      :return:      Data in list format
    '''
    with open(path, 'rb') as f:
      reader = csv.reader(f)
      data = list(reader)
    return data

  def get_range_of_data(self, data):
    '''
    Forms the occupied range of cells in the excel file, starting with A1.

    :param data:  Input data table
    :type:        list
    :return:      The range of cells that table will occupy, starting with A1
    :type:        str
    '''
    if data is None:
      raise ValueError("Invalid input data")
    row = np.shape(data)[0]
    try:
      col = np.shape(data)[1]
    except IndexError:
      raise ValueError("Invalid format of data. Number of columns in lines differ")

    degree = 26
    digits_num = 1
    while col >= degree:
      degree *= 26
      digits_num += 1
    xl_col = ''
    for i in range(digits_num, 0, -1):
      n = col // (26 ** (i - 1))
      xl_col += chr(64 + n)
      col -= n * (26 ** (i - 1))

    return "A1:{0}{1}".format(xl_col, row)

def get_port(HOST):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    while True:
        PORT = randint(5000, 6000)
        try:
            s.bind((HOST, PORT))
        except:
            pass
        else:
            s.close()
            break
    return PORT

if __name__ == "__main__":
    HOST = "localhost"
    PORT = get_port(HOST)
    excelclient = ExcelClient(RESOURCE="https://graph.microsoft.com", TENANT="your tenant",
                              AUTHORITY_HOST_URL="https://login.microsoftonline.com",
                              CLIENT_ID="your client id",
                              CLIENT_SECRET="your client secret",
                              API_VERSION="v1.0", HOST=HOST, PORT=PORT)

    data = excelclient.get_data('out.csv')
    range = excelclient.get_range_of_data(data)
    configs = excelclient.delete_range(file_id="01BEQXWXBQ2QNOPSCY4NB2EEE3V2K53RA5", sheetname="Sheet1",
                                       range=range)
    ENDPOINT = configs[0]         # The end point of the query
    TYPE_OF_REQUEST = configs[1]  # The type of the query
    REQUEST_BODY = configs[2]     # The body of the query

    @app.route("/")
    def main():
        login_url = 'http://localhost:{}/login'.format(excelclient.PORT)
        resp = flask.Response(status=307)
        resp.headers['location'] = login_url
        return resp


    @app.route("/login")
    def login():
        auth_state = str(uuid.uuid4())
        flask.session['state'] = auth_state
        authorization_url = excelclient.TEMPLATE_AUTHZ_URL.format(
            excelclient.TENANT,
            excelclient.CLIENT_ID,
            excelclient.REDIRECT_URI,
            auth_state,
            excelclient.RESOURCE)
        resp = flask.Response(status=307)
        resp.headers['location'] = authorization_url
        return resp


    @app.route("/getAToken")
    def main_logic():
        code = flask.request.args['code']
        state = flask.request.args['state']
        if state != flask.session['state']:
            raise ValueError("State does not match")
        auth_context = adal.AuthenticationContext(excelclient.AUTHORITY_URL)
        token_response = auth_context.acquire_token_with_authorization_code(code, excelclient.REDIRECT_URI, excelclient.RESOURCE,
                                                                            excelclient.CLIENT_ID, excelclient.CLIENT_SECRET)
        # It is recommended to save this to a database when using a production app.
        flask.session['access_token'] = token_response['accessToken']
        return flask.redirect('/graphcall')


    @app.route('/graphcall')
    def graphcall():
        if 'access_token' not in flask.session:
            return flask.redirect(flask.url_for('login'))
        endpoint = excelclient.RESOURCE + '/' + excelclient.API_VERSION + ENDPOINT
        http_headers = {'Authorization': 'Bearer ' + flask.session.get('access_token'),
                        'User-Agent': 'adal-python-sample',
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'client-request-id': str(uuid.uuid4())}
        if TYPE_OF_REQUEST == 'get':
            graph_data = requests.get(endpoint, headers=http_headers, stream=False).json()
        elif TYPE_OF_REQUEST == 'put':
            graph_data = requests.put(endpoint, headers=http_headers, json=REQUEST_BODY, stream=False).json()
        elif TYPE_OF_REQUEST == 'patch':
            graph_data = requests.patch(endpoint, headers=http_headers, json=REQUEST_BODY, stream=False).json()
        elif TYPE_OF_REQUEST == 'post':
            graph_data = requests.post(endpoint, headers=http_headers, json=REQUEST_BODY, stream=False).json()
        elif TYPE_OF_REQUEST == 'delete':
            graph_data = requests.delete(endpoint, headers=http_headers, json=REQUEST_BODY, stream=False).json()
        return flask.render_template('display_graph_info.html', graph_data=graph_data)

    app.run(HOST, PORT)
