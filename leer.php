<?php
	require 'Classes/PHPExcel/IOFactory.php'; //Agregamos la librería 
	//require 'conexion.php'; //Agregamos la conexión
	$mysqli=new mysqli("localhost","root","","db_codecol"); //servidor, usuario de base de datos, contraseña del usuario, nombre de base de datos
	
	if(mysqli_connect_errno()){
		echo 'Conexion Fallida : ', mysqli_connect_error();
		exit();
	}
	//Variable con el nombre del archivo
	$nombreArchivo = 'ejemplo.xlsx';
	// Cargo la hoja de cálculo
	$objPHPExcel = PHPExcel_IOFactory::load($nombreArchivo);
	
	//Asigno la hoja de calculo activa
	$objPHPExcel->setActiveSheetIndex(0);
	//Obtengo el numero de filas del archivo
	$numRows = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();
	
	echo '<table border=1><tr><td>Cedula</td><td>Nombre</td><td>Direccion</td><td>telefono</td><td>celular</td><td>FECHANAC</td><td>fecha afiliacion</td><td>direccion2</td><td>ano</td></tr>';
	
	for ($i = 1; $i <= $numRows; $i++) {
		
		$cedula = $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getCalculatedValue();
		$nombre = $objPHPExcel->getActiveSheet()->getCell('L'.$i)->getCalculatedValue();
		$direccion = $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getCalculatedValue();
		$telefono = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getCalculatedValue();
		$celular = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getCalculatedValue();
		$fechanac = $objPHPExcel->getActiveSheet()->getCell('M'.$i)->getCalculatedValue();
		$fechaafi = $objPHPExcel->getActiveSheet()->getCell('O'.$i)->getCalculatedValue();
		$direccion2 = $objPHPExcel->getActiveSheet()->getCell('B'.$i)->getCalculatedValue();
		$ano = $objPHPExcel->getActiveSheet()->getCell('P'.$i)->getCalculatedValue();
		$fechanacimiento = date('Y/m/d', PHPExcel_Shared_Date::ExcelToPHP($fechanac));
		$fechaafiliacion = date('Y/m/d', PHPExcel_Shared_Date::ExcelToPHP($fechaafi));
		$anio_aportacion = number_format($ano,2);
		echo '<tr>';
		echo '<td>'. $cedula.'</td>';
		echo '<td>'. $nombre.'</td>';
		echo '<td>'. $direccion.'</td>';
		echo '<td>'. $telefono.'</td>';
		echo '<td>'. $celular.'</td>';
		echo '<td>'. $fechanacimiento.'</td>';
		echo '<td>'. $fechaafiliacion.'</td>';
		echo '<td>'. $direccion2.'</td>';
		echo '<td>'. $anio_aportacion.'</td>';
		echo '</tr>';
		
		$sql = "INSERT INTO `billing_cliente` (`billing_cliente_id`, `es_pasaporte`, `PersonaComercio_cedulaRuc`, `nombres`, `apellidos`, `razonsocial`, `nombre_comercial`, `direccion`, `diasCredito`, `pais`, `ciudad`, `comentarios`, `clientetipo_idclientetipo`, `descuentomaxporcent`, `cupocredito`, `email`, `telefonos`, `celular`, `docidentificacion_id`, `vendedor_id`, `fecha`, `usuario`, `clave`, `cupo_temporal`, `tipo_ruc`, `descuentotemp`, `clase`, `provincia`, `canton`, `parroquia`, `sexo`, `estado_civil`, `origen_ingresos`, `tipo_identificacion`, `aseguradora_id`, `cuenta_gasto`, `credito`, `id_sector`, `estaActivo`, `id_nro_poste`, `codigo_cliente`, `descuento_valor`, `edad_cli`, `fecha_nacimiento_cli`, `profesion_cli`, `es_parking`, `imagen`, `fecha_creacion_cli`, `direccion2`, `anio_aportacion`, `fecha_fallecimiento`, `celular2`, `categoria_id`) VALUES ('0', '0', '$cedula', '$nombre', '.', NULL, NULL, '$direccion', '0', NULL, 'Loja', NULL, '13', '0', '0', NULL, '$telefono', '$celular', NULL, NULL, NULL, NULL, NULL, '0.00000', '', '0.00', NULL, NULL, NULL, NULL, 'Null', NULL, NULL, NULL, '-2', NULL, '0', '0', '1', NULL, NULL, NULL, NULL, '$fechanacimiento', NULL, '0', NULL, '$fechaafiliacion', '$direccion2', '$anio_aportacion', NULL, NULL, NULL);";
		$result = $mysqli->query($sql);
	}
	
	echo '<table>';
?>