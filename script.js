let workbookModelo;
fetch("RDO_MODELO.xlsx")
  .then(res => { if (!res.ok) throw new Error("Falha ao buscar o arquivo modelo"); return res.arrayBuffer(); })
  .then(data => { workbookModelo = XLSX.read(data, { type: "array" }); console.log("Modelo XLSX carregado com sucesso!"); })
  .catch(err => { console.error("Erro ao carregar o modelo:", err); alert("Não foi possível carregar o modelo RDO_MODELO.XLSX."); });

function abrirDialogGerar(){document.getElementById("dialog-gerar").style.display="block";}
function abrirDialogLimpar(){document.getElementById("dialog-limpar").style.display="block";}
function fecharDialog(id){document.getElementById(id).style.display="none";}

function gerarRDO(){
  if(!workbookModelo){alert("O modelo ainda não foi carregado.");return;}
  const sheet=workbookModelo.Sheets["RDO"];
  const data=document.getElementById("data").value;
  const sigla=document.getElementById("siglaEquipe").value;
  const dataFormatada=data.split("-").reverse().join("/");
  sheet["E6"]={t:"s",v:dataFormatada};
  sheet["F35"]={t:"s",v:sigla};
  sheet["AD14"]={t:"n",v:parseInt(document.getElementById("hist_encarregado").value)||""};
  sheet["AD15"]={t:"n",v:parseInt(document.getElementById("hist_motorista").value)||""};
  const kmInicial=document.getElementById("km_inicial_1").value;
  const kmFinal=document.getElementById("km_final_1").value;
  const horarioInicio=document.getElementById("hora_inicio_1").value;
  const horarioFim=document.getElementById("hora_fim_1").value;
  const ordemServico=document.getElementById("ordem_servico_1").value;
  sheet["F37"]={t:"s",v:`Km ${kmInicial} a ${kmFinal}`};
  sheet["F41"]={t:"s",v:`Horário de Início: ${horarioInicio}`};
  sheet["F42"]={t:"s",v:`Horário de Término: ${horarioFim}`};
  sheet["F40"]={t:"s",v:`Ordem de serviço: ${ordemServico}`};
  const nomeArquivo=`${sigla}_${data.split("-").reverse().join("-")}.xlsx`;
  const wbout=XLSX.write(workbookModelo,{bookType:"xlsx",type:"array"});
  saveAs(new Blob([wbout],{type:"application/octet-stream"}),nomeArquivo);
}
function executarGeracao(){fecharDialog("dialog-gerar");gerarRDO();alert("RDO gerado com sucesso!");window.scrollTo({top:0,behavior:"smooth"});}
function executarLimpeza(){fecharDialog("dialog-limpar");document.querySelectorAll("input,select").forEach(el=>{if(el.type==="checkbox"||el.type==="radio")el.checked=false;else el.value="";});alert("Formulário limpo.");}