
class SQLDb
    attr_accessor :mdb, :connection, :data, :fields

    def initialize(mdb=nil)
        @mdb = mdb
        @connection = nil
        @data = nil
        @fields = nil
    end

    def open
        connection_string = 'Driver={SQL Server Native Client 10.0};Server=.\sqlexpress;Trusted_Connection=yes;Database='
        connection_string << @mdb
        @connection = WIN32OLE.new('ADODB.Connection')
        @connection.Open(connection_string)
    end

    def query(sql)
        recordset = WIN32OLE.new('ADODB.Recordset')
        recordset.Open(sql, @connection)
        @fields = []
        recordset.Fields.each do |field|
            @fields << field.Name
        end
        begin
            @data = recordset.GetRows.transpose
        rescue
            @data = []
        end
        recordset.Close
    end

    def execute(sql)
        @connection.Execute(sql)
    end

    def close
        @connection.Close
    end
end
