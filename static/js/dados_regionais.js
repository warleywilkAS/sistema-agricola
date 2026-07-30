// Gerado automaticamente a partir da lista de Regionais/Municipios/Mesorregioes
// fornecida para o Sistema de Coleta de Dados Agricolas - IDR-Parana

const REGIOES_ORDEM = [
  "Apucarana",
  "Campo Mourão",
  "Cascavel",
  "Cianorte",
  "Cornélio Procópio",
  "Curitiba",
  "Dois Vizinhos",
  "Francisco Beltrão",
  "Guarapuava",
  "Irati",
  "Ivaiporã",
  "Laranjeiras do Sul",
  "Londrina",
  "Maringá",
  "Paranaguá",
  "Paranavaí",
  "Pato Branco",
  "Ponta Grossa",
  "Sto Antônio da Platina",
  "Toledo",
  "Umuarama",
  "União da Vitória"
];

const MESO_POR_REGIAO = {
  "Apucarana": "Norte",
  "Campo Mourão": "Noroeste",
  "Cascavel": "Oeste",
  "Cianorte": "Noroeste",
  "Cornélio Procópio": "Norte",
  "Curitiba": "Metropolitana e Litoral",
  "Dois Vizinhos": "Sudoeste",
  "Francisco Beltrão": "Sudoeste",
  "Guarapuava": "Centro Sul",
  "Irati": "Centro Sul",
  "Ivaiporã": "Norte",
  "Laranjeiras do Sul": "Centro Sul",
  "Londrina": "Norte",
  "Maringá": "Noroeste",
  "Paranaguá": "Metropolitana e Litoral",
  "Paranavaí": "Noroeste",
  "Pato Branco": "Sudoeste",
  "Ponta Grossa": "Centro",
  "Sto Antônio da Platina": "Norte",
  "Toledo": "Oeste",
  "Umuarama": "Noroeste",
  "União da Vitória": "Centro Sul"
};

const MUNICIPIOS_POR_REGIAO = {
  "Apucarana": [
    "Apucarana",
    "Arapongas",
    "Bom Sucesso",
    "Califórnia",
    "Cambira",
    "Jandaia do Sul",
    "Kaloré",
    "Marilândia do Sul",
    "Marumbi",
    "Mauá da Serra",
    "Novo Itacolomi",
    "Rio Bom",
    "Sabáudia"
  ],
  "Campo Mourão": [
    "Altamira do Paraná",
    "Araruna",
    "Barbosa Ferraz",
    "Boa Esperança",
    "Campina da Lagoa",
    "Campo Mourão",
    "Corumbataí do Sul",
    "Engenheiro Beltrão",
    "Farol",
    "Fênix",
    "Goioerê",
    "Iretama",
    "Janiópolis",
    "Luiziana",
    "Mamborê",
    "Moreira Sales",
    "Nova Cantu",
    "Peabiru",
    "Quarto Centenário",
    "Quinta do Sol",
    "Rancho Alegre do Oeste",
    "Roncador",
    "Terra Boa",
    "Ubiratã"
  ],
  "Cascavel": [
    "Anahy",
    "Boa Vista da Aparecida",
    "Braganey",
    "Cafelândia",
    "Campo Bonito",
    "Capitão Leônidas Marques",
    "Cascavel",
    "Catanduvas",
    "Corbélia",
    "Céu Azul",
    "Diamante do Oeste",
    "Foz do Iguaçu",
    "Ibema",
    "Itaipulândia",
    "Lindoeste",
    "Matelândia",
    "Medianeira",
    "Missal",
    "Nova Aurora",
    "Ramilândia",
    "Santa Lúcia",
    "Santa Tereza do Oeste",
    "Santa Terezinha do Itaipu",
    "Serranópolis do Iguaçu",
    "São Miguel do Iguaçu",
    "Três Barras do Paraná",
    "Vera Cruz do Oeste"
  ],
  "Cianorte": [
    "Cianorte",
    "Cidade Gaúcha",
    "Guaporema",
    "Indianópolis",
    "Japurá",
    "Jussara",
    "Rondon",
    "São Manoel do Paraná",
    "São Tomé",
    "Tapejara",
    "Tuneiras do Oeste"
  ],
  "Cornélio Procópio": [
    "Abatiá",
    "Andirá",
    "Assaí",
    "Bandeirantes",
    "Congonhinhas",
    "Cornélio Procópio",
    "Itambaracá",
    "Jataizinho",
    "Leópolis",
    "Nova América da Colina",
    "Nova Fátima",
    "Nova Santa Bárbara",
    "Rancho Alegre",
    "Santa Amélia",
    "Santa Cecília do Pavão",
    "Santa Mariana",
    "Santo Antônio do Paraíso",
    "Sapopema",
    "Sertaneja",
    "São Jerônimo da Serra",
    "São Sebastião da Amoreira",
    "Uraí"
  ],
  "Curitiba": [
    "Adrianópolis",
    "Agudos do Sul",
    "Almirante Tamandaré",
    "Araucária",
    "Balsa Nova",
    "Bocaiuva do Sul",
    "Campina Grande do Sul",
    "Campo Largo",
    "Campo Magro",
    "Campo do Tenente",
    "Cerro Azul",
    "Colombo",
    "Contenda",
    "Doutor Ulysses",
    "Fazenda Rio Grande",
    "Itaperuçu",
    "Lapa",
    "Mandirituba",
    "Pinhais",
    "Piraquara",
    "Piên",
    "Quatro Barras",
    "Quitandinha",
    "Rio Branco do Sul",
    "Rio Negro",
    "São José dos Pinhais",
    "Tijucas do Sul",
    "Tunas do Paraná"
  ],
  "Dois Vizinhos": [
    "Boa Esperança do Iguaçu",
    "Cruzeiro do Iguaçu",
    "Dois Vizinhos",
    "Nova Esperança do Sudoeste",
    "Nova Prata do Iguaçu",
    "Salto do Lontra",
    "São Jorge D'Oeste"
  ],
  "Francisco Beltrão": [
    "Ampére",
    "Barracão",
    "Bela Vista da Caroba",
    "Bom Jesus do Sul",
    "Capanema",
    "Enéas Marques",
    "Flor da Serra do Sul",
    "Francisco Beltrão",
    "Manfrinópolis",
    "Marmeleiro",
    "Pinhal de São Bento",
    "Planalto",
    "Pérola do Oeste",
    "Realeza",
    "Renascença",
    "Salgado Filho",
    "Santa Izabel do Oeste",
    "Santo Antônio do Sudoeste",
    "Verê"
  ],
  "Guarapuava": [
    "Boa Ventura de São Roque",
    "Campina do Simão",
    "Candói",
    "Cantagalo",
    "Foz do Jordão",
    "Goioxim",
    "Guarapuava",
    "Laranjal",
    "Mato Rico",
    "Palmital",
    "Pinhão",
    "Pitanga",
    "Prudentópolis",
    "Santa Maria do Oeste",
    "Turvo"
  ],
  "Irati": [
    "Fernandes Pinheiro",
    "Guamiranga",
    "Imbituva",
    "Inácio Martins",
    "Irati",
    "Mallet",
    "Rebouças",
    "Rio Azul",
    "Teixeira Soares"
  ],
  "Ivaiporã": [
    "Arapuã",
    "Ariranha do Ivaí",
    "Borrazópolis",
    "Cruzmaltina",
    "Cândido de Abreu",
    "Faxinal",
    "Godoy Moreira",
    "Grandes Rios",
    "Ivaiporã",
    "Jardim Alegre",
    "Lidianópolis",
    "Lunardelli",
    "Manoel Ribas",
    "Rio Branco do Ivaí",
    "Rosário do Ivaí",
    "São João do Ivaí",
    "São Pedro do Ivaí"
  ],
  "Laranjeiras do Sul": [
    "Diamante do Sul",
    "Espigão Alto do Iguaçu",
    "Guaraniaçu",
    "Laranjeiras do Sul",
    "Marquinho",
    "Nova Laranjeiras",
    "Porto Barreiro",
    "Quedas do Iguaçu",
    "Rio Bonito do Iguaçu",
    "Virmond"
  ],
  "Londrina": [
    "Alvorada do Sul",
    "Bela Vista do Paraíso",
    "Cafeara",
    "Cambé",
    "Centenário do Sul",
    "Florestópolis",
    "Guaraci",
    "Ibiporã",
    "Jaguapitã",
    "Londrina",
    "Lupionópolis",
    "Miraselva",
    "Pitangueiras",
    "Prado Ferreira",
    "Primeiro de Maio",
    "Rolândia",
    "Sertanópolis",
    "Tamarana"
  ],
  "Maringá": [
    "Astorga",
    "Atalaia",
    "Colorado",
    "Doutor Camargo",
    "Floraí",
    "Floresta",
    "Flórida",
    "Iguaraçu",
    "Itambé",
    "Ivatuba",
    "Lobato",
    "Mandaguaçu",
    "Marialva",
    "Maringá",
    "Munhoz de Mello",
    "Nossa Senhora das Graças",
    "Nova Esperança",
    "Ourizona",
    "Paiçandu",
    "Presidente Castelo Branco",
    "Santa Fé",
    "Santa Inês",
    "Santo Inácio",
    "Sarandi",
    "São Jorge do Ivaí",
    "Uniflor",
    "Ângulo"
  ],
  "Paranaguá": [
    "Antonina",
    "Guaraqueçaba",
    "Guaratuba",
    "Matinhos",
    "Morretes",
    "Paranaguá",
    "Pontal do Paraná"
  ],
  "Paranavaí": [
    "Alto Paraná",
    "Amaporã",
    "Cruzeiro do Sul",
    "Diamante do Norte",
    "Guairaçá",
    "Inajá",
    "Itaúna do Sul",
    "Jardim Olinda",
    "Loanda",
    "Marilena",
    "Mirador",
    "Nova Aliança do Ivaí",
    "Nova Londrina",
    "Paranapoema",
    "Paranavaí",
    "Paraíso do Norte",
    "Planaltina do Paraná",
    "Porto Rico",
    "Querência do Norte",
    "Santa Cruz do Monte Castelo",
    "Santa Isabel do Ivaí",
    "Santa Mônica",
    "Santo Antônio do Caiuá",
    "São Carlos do Ivaí",
    "São João do Caiuá",
    "São Pedro do Paraná",
    "Tamboara",
    "Terra Rica"
  ],
  "Pato Branco": [
    "Bom Sucesso do Sul",
    "Chopinzinho",
    "Clevelândia",
    "Coronel Domingos Soares",
    "Coronel Vivida",
    "Honório Serpa",
    "Itapejara do Oeste",
    "Mangueirinha",
    "Mariópolis",
    "Palmas",
    "Pato Branco",
    "Saudades do Iguaçu",
    "São João",
    "Vitorino"
  ],
  "Ponta Grossa": [
    "Arapoti",
    "Carambeí",
    "Castro",
    "Imbaú",
    "Ipiranga",
    "Ivaí",
    "Jaguariaíva",
    "Ortigueira",
    "Palmeira",
    "Piraí do Sul",
    "Ponta Grossa",
    "Porto Amazonas",
    "Sengés",
    "São João do Triunfo",
    "Telêmaco Borba",
    "Tibagi",
    "Ventania"
  ],
  "Sto Antônio da Platina": [
    "Barra do Jacaré",
    "Cambará",
    "Carlópolis",
    "Conselheiro Mairinck",
    "Curiúva",
    "Figueira",
    "Guapirama",
    "Ibaiti",
    "Jaboti",
    "Jacarezinho",
    "Japira",
    "Joaquim Távora",
    "Pinhalão",
    "Quatiguá",
    "Ribeirão Claro",
    "Salto do Itararé",
    "Santana do Itararé",
    "Santo Antônio da Platina",
    "Siqueira Campos",
    "São José da Boa Vista",
    "Tomazina",
    "Wenceslau Braz"
  ],
  "Toledo": [
    "Assis Chateaubriand",
    "Entre Rios do Oeste",
    "Formosa do Oeste",
    "Guaíra",
    "Iracema do Oeste",
    "Jesuítas",
    "Marechal Cândido Rondon",
    "Maripá",
    "Mercedes",
    "Nova Santa Rosa",
    "Ouro Verde do Oeste",
    "Palotina",
    "Quatro Pontes",
    "Santa Helena",
    "São José das Palmeiras",
    "São Pedro do Iguaçu",
    "Terra Roxa",
    "Toledo",
    "Tupãssi"
  ],
  "Umuarama": [
    "Alto Paraíso",
    "Alto Piquiri",
    "Altônia",
    "Brasilândia do Sul",
    "Cafezal do Sul",
    "Cruzeiro do Oeste",
    "Douradina",
    "Esperança Nova",
    "Francisco Alves",
    "Icaraíma",
    "Iporã",
    "Ivaté",
    "Mariluz",
    "Nova Olímpia",
    "Perobal",
    "Pérola",
    "São Jorge do Patrocínio",
    "Tapira",
    "Umuarama",
    "Xambrê"
  ],
  "União da Vitória": [
    "Antônio Olinto",
    "Bituruna",
    "Cruz Machado",
    "General Carneiro",
    "Paula Freitas",
    "Paulo Frontin",
    "Porto Vitória",
    "São Mateus do Sul",
    "União da Vitória"
  ]
};


/**
 * Preenche o <select> de Regiao com as opcoes fixas, e liga o evento de
 * troca para atualizar automaticamente Municipio (lista filtrada) e
 * MESO_IDR (preenchido sozinho, somente leitura).
 *
 * @param {string} idRegiao   id do <select> de Regiao
 * @param {string} idMunicipio id do <select> de Municipio
 * @param {string} idMeso     id do campo (input) de MESO_IDR
 * @param {string} [valorRegiaoAtual]    valor ja selecionado (tela de edicao)
 * @param {string} [valorMunicipioAtual] valor ja selecionado (tela de edicao)
 */
function initRegiaoCascata(idRegiao, idMunicipio, idMeso, valorRegiaoAtual, valorMunicipioAtual) {
    const selRegiao = document.getElementById(idRegiao);
    const selMunicipio = document.getElementById(idMunicipio);
    const campoMeso = document.getElementById(idMeso);

    if (!selRegiao || !selMunicipio) return;

    // popula as opcoes de Regiao (uma vez)
    if (selRegiao.options.length <= 1) {
        REGIOES_ORDEM.forEach(function (regiao) {
            const opt = document.createElement("option");
            opt.value = regiao;
            opt.textContent = regiao;
            selRegiao.appendChild(opt);
        });
    }

    function atualizaMunicipios(regiaoSelecionada, municipioParaSelecionar) {
        selMunicipio.innerHTML = "";
        const optVazia = document.createElement("option");
        optVazia.value = "";
        optVazia.textContent = "Selecione...";
        selMunicipio.appendChild(optVazia);

        const lista = MUNICIPIOS_POR_REGIAO[regiaoSelecionada] || [];
        lista.forEach(function (municipio) {
            const opt = document.createElement("option");
            opt.value = municipio;
            opt.textContent = municipio;
            if (municipioParaSelecionar && municipio === municipioParaSelecionar) {
                opt.selected = true;
            }
            selMunicipio.appendChild(opt);
        });

        if (campoMeso) {
            campoMeso.value = MESO_POR_REGIAO[regiaoSelecionada] || "";
        }
    }

    selRegiao.addEventListener("change", function () {
        atualizaMunicipios(selRegiao.value, null);
    });

    // estado inicial (tela de edicao: regiao/municipio ja preenchidos)
    if (valorRegiaoAtual) {
        selRegiao.value = valorRegiaoAtual;
        atualizaMunicipios(valorRegiaoAtual, valorMunicipioAtual);
    }
}
