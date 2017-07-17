# !/usr/bin/python

from ncclient import manager
from ncclient.transport import errors
import sys, time, telnetlib
from xlrd import open_workbook
from tempfile import TemporaryFile
from xlwt import Workbook, easyxf
from xlwt import Style

global conn, sessionId
global operations, dataStores
global filterData, configData, book, clicommandData, outputfilterData
global row_count_to_append_result, active_sheet_name, adding_sheet

row_count_to_append_result = 0
active_sheet_name = ""

book = open_workbook("C:\Users\ss015282\Box Sync\PycharmProjects\Github\/ncclient\RPC_XML_Data.xlsx")
write_to_book = Workbook()

def connect(host, port, user, password):
    global conn, sessionId
    global operations, dataStores
    global filterData, configData, book, clicommandData, outputfilterData
    global row_count_to_append_result, active_sheet_name

    """
        filterData = '''
    						<if:interfaces xmlns:if="urn:ietf:params:xml:ns:yang:ietf-interfaces">
    							<if:interface>
    								<if:name>0/1</if:name>
    								<if:description/>
    							</if:interface>
    						</if:interfaces>
    				 '''
        configData = '''
    					<nc:config xmlns:nc="urn:ietf:params:xml:ns:netconf:base:1.0">
    						<interfaces xmlns="urn:ietf:params:xml:ns:yang:ietf-interfaces">
    							<interface>
    								<name>0/1</name>
    								<description nc:operation="%s">%s</description>
    							</interface>
    						</interfaces>
    					</nc:config>
    				 '''
    """

    try:
        # connect to the Netconf server
        conn = manager.connect(host=host, port=port, username=user, password=password, hostkey_verify=False)

        sessionId = conn.session_id

        print 'connected:', conn.connected, ' .... to host', host, 'on port:', port

        # Get session parameters
        print 'session-id:', sessionId
        print 'client capabilities:'
        print '####################'
        for i in conn.client_capabilities:
            print ' ', i
        print 'server capabilities:'
        print '####################'
        for i in conn.server_capabilities:
            print ' ', i
        # return this conn object to process edit and get operations to __main__ method
        return conn

    except errors.SSHError:
        print 'Unable to connect to host:', host, 'on port:', port

# Method will be called while locking the data store
def datastore_lock(datastore):
    print '\nlocking the datastore :' + datastore
    conn.lock(datastore)

# Method will be called while unlocking the data store
def datastore_unlock(datastore):
    print '\nunlocking the datastore :' + datastore
    conn.unlock(datastore)

def get_config_intf_description(datastore, filterData, containername):
    print 'Retrieving config using filter from %s datastore for %s container, please wait ...' % (datastore, containername)
    print '\n Request filterData :' + '\n' + filterData
    get_config_response = conn.get_config(source=datastore, filter=('subtree', filterData)).data_xml
    return get_config_response


def edit_config_intf_description():
    operations = ["merge", "remove", "replace", "delete", "create"]
    dataStores = ["running", "startup", "candidate"]

    for sheet_index in range(book.nsheets):
        sheet_index_number = book.sheet_by_index(sheet_index)
        # Initializing the number to 0. This value will be used to append the output date to Excel rows.
        row_count_to_append_result = 0

        for row in range(1, sheet_index_number.nrows):
            node_name = sheet_index_number.row(row)[0].value
            filterData = sheet_index_number.row(row)[1].value
            configData = sheet_index_number.row(row)[2].value
            clioutputData = sheet_index_number.row(row)[3].value
            clicommandData = sheet_index_number.row(row)[4].value

            for datastore in dataStores:
                try:
                    # Lock the datastore until we finish using that datastore
                    datastore_lock(datastore)
                    # we have to make sure both startup and candidate datastores are having proper fields before working on them
                    # So copying from running (which always have proper fields) to startup and candidate and working on them
                    if datastore == "startup":
                        conn.copy_config("running", "startup")
                    if datastore == "candidate":
                        conn.copy_config("running", "candidate")
                    for operation in operations:
                        try:
                            # Starting from row count 1. because row value 0 is left for headers.
                            row_count_to_append_result = row_count_to_append_result + 1
                            # Perform Edit operation based on datastore and operation
                            dataConfig = configData % operation
                            print '###################################################################'
                            print "\n edit_config" + " operation: " + operation + " on datastore: " + datastore + " on container: " + sheet_index_number.name
                            print '\n Request configData :' + '\n' + dataConfig
                            edit_config_response = conn.edit_config(target=datastore, config=dataConfig)
                            print '\n Response from server for configData :' + '\n' + str(edit_config_response)
                            time.sleep(2)
                            # Performing Get-config operation to check the edit-config data was successfully configured or not
                            print "\n get_config" + " after operation: " + operation + " on datastore: " + datastore + " on container: " + sheet_index_number.name
                            get_config_response_output = get_config_intf_description(datastore, filterData, sheet_index_number.name)
                            print '\n Response from server for filterData :' + '\n' + get_config_response_output
                            print '###################################################################'
                            time.sleep(2)

                            # telnet cli output check is required when datastore is running
                            if datastore != "running":
                                # send the data to form into excel file
                                write_results_to_sheet(operation, datastore, node_name, sheet_index,
                                                       sheet_index_number.name, row_count_to_append_result, dataConfig,
                                                       edit_config_response, filterData, get_config_response_output,
                                                       clicommandData="None", telnet_cli_output="None")
                            else:
                                telnet_cli_output = telnet_dut(clicommandData)

                                # send the data to form into excel file
                                write_results_to_sheet(operation, datastore, node_name, sheet_index,
                                                       sheet_index_number.name, row_count_to_append_result, dataConfig,
                                                       edit_config_response, filterData, get_config_response_output,
                                                       clicommandData, telnet_cli_output)

                        except errors.NCClientError as e:
                            print '\n Response from server :' + '\n' + str(e.message)
                            write_results_to_sheet(operation, datastore, node_name, sheet_index, sheet_index_number.name, row_count_to_append_result,\
                                                   dataConfig, str(e.message), filterData = "None",\
                                                   get_config_response_output = "None", clicommandData = "None", telnet_cli_output = "None")

                            # Eventhough exceptioin caugnt we wanted to rotate the for loop to continue with our requests
                            pass
                    # unlock the datastore after doing all the operations
                    datastore_unlock(datastore)

                except errors.NCClientError as e:
                    print e.message
                    pass

def write_results_to_sheet(operation, datastore, node_name, sheetnum, sheetname, row_count_to_append_result, dataConfig, edit_config_response, filterData, get_config_response_output, clicommandData, telnet_cli_output):
    try:
        adding_sheet = write_to_book.add_sheet(sheetname, cell_overwrite_ok=True)

        adding_sheet.write(0, 0, "DataStore",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))
        adding_sheet.write(0, 1, "Node",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))
        adding_sheet.write(0, 2, "Operation",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))
        adding_sheet.write(0, 3, "dataConfig_Request_XML",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))
        adding_sheet.write(0, 4, "dataConfig_Response_From_Server",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))
        adding_sheet.write(0, 5, "filterData_Request_XML",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))
        adding_sheet.write(0, 6, "filterData_Response_From_Server",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))
        adding_sheet.write(0, 7, "clicommandData",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))
        adding_sheet.write(0, 8, "telnetCliOutput",
                         Style.easyxf('pattern: pattern solid, fore_colour green;' 'borders: left thick, right thick, top thick, bottom thick;' 'font:height 500;' 'align: wrap yes'))

        style = easyxf('borders: left thin, right thin, top thin, bottom thin;' 'align: wrap yes')

        adding_sheet.col(int(row_count_to_append_result) - 1).width = 15000
        adding_sheet.row(int(row_count_to_append_result)).height_mismatch = 1
        adding_sheet.row(int(row_count_to_append_result)).height = 3000

        row = adding_sheet.row(row_count_to_append_result)
        row.write(0, datastore, style)
        row.write(1, node_name, style)
        row.write(2, operation, style)
        row.write(3, dataConfig, style)
        row.write(4, str(edit_config_response), style)
        row.write(5, filterData, style)
        row.write(6, str(get_config_response_output), style)
        row.write(7, clicommandData, style)
        row.write(8, telnet_cli_output, style)

    # Here Exception will be caught when "write_to_book.add_sheet" is trying to add already existing sheet.
    # So instead of adding it, we are calling that existing sheet to performing our operations.
    except:
        adding_sheet = write_to_book.get_sheet(sheetnum)

        style = easyxf('borders: left thin, right thin, top thin, bottom thin;' 'align: wrap yes')

        adding_sheet.col(int(row_count_to_append_result) - 1).width = 15000
        adding_sheet.row(int(row_count_to_append_result)).height_mismatch = 1
        adding_sheet.row(int(row_count_to_append_result)).height = 3000

        row = adding_sheet.row(row_count_to_append_result)
        row.write(0, datastore, style)
        row.write(1, node_name, style)
        row.write(2, operation, style)
        row.write(3, dataConfig, style)
        row.write(4, str(edit_config_response), style)
        row.write(5, filterData, style)
        row.write(6, str(get_config_response_output), style)
        row.write(7, clicommandData, style)
        row.write(8, telnet_cli_output, style)

    write_to_book.save("C:\Users\ss015282\Box Sync\PycharmProjects\Github\/ncclient\RPC_XML_Data_Test_Results.xls")
    write_to_book.save(TemporaryFile())

def telnet_dut(clicommandData):
    tn = telnetlib.Telnet("10.130.170.252")
    # tn.read_until(" Entering server port, ..... type ^z for port menu.")
    # tn.write("")
    tn.read_until("User:")
    tn.write("admin\n")
    tn.read_until("Password:")
    tn.write("\n")
    tn.write("enable\n")
    tn.read_until("#")
    tn.write("terminal length 0\n")
    tn.read_until("#")
    tn.write(str(clicommandData)+"\n")
    clicommmandOutput = tn.read_until("#")
    print '#############################'
    print "CLI command : %s" % clicommandData
    print "\nCLI output :", clicommmandOutput
    print '#############################'
    return clicommmandOutput

if __name__ == '__main__':
    conn = connect("10.130.170.252", 830, "admin", "")
    print "\nGoing to perform Netconf Edit-conifg and Get-config operations......!!!"
    if conn.connected:
        edit_config_intf_description()

    # Make sure session is closed after finishing work
    print "Closing the connection to Server...!!!"
    conn.close_session()
