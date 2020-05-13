Sub Main
	Call RelateDatabase()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Archivo: Conector visual
Function RelateDatabase
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.VisualConnector
	id0 = task.AddDatabase ("Ejemplo-Detalle de ventas.IMD")
	id1 = task.AddDatabase ("Ejemplo-Vendedores.IMD")
	id2 = task.AddDatabase ("Ejemplo-Clientes.IMD")
	task.MasterDatabase = id0
	task.AppendDatabaseNames = FALSE
	task.IncludeAllPrimaryRecords = TRUE
	task.AddRelation id0, "NUM_VENDEDOR", id1, "NUM_VENDEDOR"
	task.AddRelation id0, "NUM_CLI", id2, "NUM_CLI"
	task.IncludeAllFields
	task.CreateVirtualDatabase = False
	dbName = "Detalle_total01.IMD"
	task.OutputDatabaseName = dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function