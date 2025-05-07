import uiautomation as auto

# Obtener todas las ventanas abiertas
for w in auto.GetRootControl().GetChildren():
    print(f"Ventana: {w.Name}, Clase: {w.ClassName}")

sap_window = auto.WindowControl(Name="Evaluación del log de auditoría de seguridad")
if sap_window.Exists():
    print("Ventana SAP encontrada")
else:
    print("No se encontró la ventana")

#for elemento in sap_window.GetChildren():
#        print(f"🔹 {elemento.ControlTypeName} - {elemento.Name} (AutoID: {elemento.AutomationId})")

for w in sap_window.GetChildren():
    print(f"Control: {w.Name}, AutomationId: {w.AutomationId}, ControlType: {w.ControlTypeName}")
    
#treelist = sap_window.PaneControl(Name="SAP's Advanced Treelist")
#treelist = sap_window.PaneControl(AutomationId="198515408")

#"192413296"
#ClassName:	"SAPALVGridControl"

pane_control = sap_window.PaneControl(AutomationId="192413296")
if pane_control.Exists():
    print("Contenedor encontrado, explorando sus hijos...")
    for child in pane_control.GetChildren():
        print(f"Elemento: {child.Name}, ControlType: {child.ControlTypeName}, AutomationId: {child.AutomationId}")
else:
    print("No se encontró el contenedor con AutomationId 192413296")
