// === 1. ESTADO GLOBAL DO SISTEMA ===

// --- CONFIGURAÇÃO DO FIREBASE ---
const firebaseConfig = {
  apiKey: "SUA_CHAVE_AQUI",
  authDomain: "seu-projeto.firebaseapp.com",
  projectId: "seu-projeto",
  storageBucket: "seu-projeto.appspot.com",
  messagingSenderId: "1234567890",
  appId: "1:1234567890:web:abcdef"
};

// Inicializa o Firebase
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore(); 
let estoque = [];
let movimentacoes = [];
let notasSolicitacao = [];
let usuariosRegistrados = [];
let usuarioLogado = null;
let categoriaAtual = 'Todas'; // Para o filtro

const dadosIniciais = [
    { codigo: "1010", nome: "Broca de Vídea 10mm", categoria: "Ferramentas", unidade: "UN", quantidade: 3, estoqueMinimo: 5, valor: 25.50, localizacao: "A-01" },
    { codigo: "2020", nome: "Chave Fixa 13mm", categoria: "Ferramentas", unidade: "UN", quantidade: 8, estoqueMinimo: 5, valor: 42.00, localizacao: "B-03" },
    { codigo: "3030", nome: "Fita Isolante 20m WEG", categoria: "Elétrica", unidade: "RL", quantidade: 45, estoqueMinimo: 20, valor: 12.90, localizacao: "C-02" },
    { codigo: "4040", nome: "Disjuntor Bifásico 32A", categoria: "Elétrica", unidade: "UN", quantidade: 15, estoqueMinimo: 10, valor: 68.00, localizacao: "D-01" },
    { codigo: "5050", nome: "Capacete de Segurança", categoria: "EPI", unidade: "UN", quantidade: 2, estoqueMinimo: 10, valor: 35.00, localizacao: "E-05" }
];

function showToast(mensagem) {
    let container = document.getElementById('toast-container');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toast-container';
        document.body.appendChild(container);
    }
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerText = mensagem;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

function setButtonLoading(btn, isLoading) {
    if (!btn) return;
    if (isLoading) {
        btn.dataset.originalText = btn.innerText;
        btn.innerText = "Processando...";
        btn.disabled = true;
    } else {
        btn.innerText = btn.dataset.originalText;
        btn.disabled = false;
    }
}

// === 2. INICIALIZAÇÃO DO SISTEMA ===
function iniciarSistema() {
    estoque = JSON.parse(localStorage.getItem('estoque'));
    if (!estoque) {
        estoque = dadosIniciais; // Carrega os dados iniciais se for a 1ª vez
        localStorage.setItem('estoque', JSON.stringify(estoque));
    }
    movimentacoes = JSON.parse(localStorage.getItem('movimentacoes')) || [];
    notasSolicitacao = JSON.parse(localStorage.getItem('notasSolicitacao')) || [];
    usuariosRegistrados = JSON.parse(localStorage.getItem('usuariosRegistrados')) || [];

    document.addEventListener('DOMContentLoaded', () => {
        configurarEventosDoSistema();
        document.getElementById('loginScreen')?.classList.add('active');
    });
}

// === 3. EVENTOS GERAIS E NAVEGAÇÃO ===
function configurarEventosDoSistema() {
    
    // Menu Responsivo
    const menuBtn = document.getElementById('menuHamburgerBtn');
    const closeBtn = document.getElementById('closeSidebarBtn');
    const sidebar = document.getElementById('sidebar');

    if (menuBtn && sidebar) { menuBtn.onclick = (e) => { e.preventDefault(); sidebar.classList.add('open'); }; }
    if (closeBtn && sidebar) { closeBtn.onclick = (e) => { e.preventDefault(); sidebar.classList.remove('open'); }; }

    // Toggle Dark Mode
    const darkBtn = document.getElementById('toggleDarkModeBtn');
    if (darkBtn) {
        darkBtn.onclick = (e) => {
            e.preventDefault();
            document.body.classList.toggle('dark-mode');
            const isDark = document.body.classList.contains('dark-mode');
            localStorage.setItem('darkMode', isDark);
        };
    }

    // Alternância de Abas
    document.querySelectorAll('.sidebar-link').forEach(link => {
        link.onclick = function(e) {
            const tabId = this.getAttribute('data-tab');
            if (!tabId) return; 
            e.preventDefault();
            
            document.querySelectorAll('.sidebar-link').forEach(l => l.classList.remove('active'));
            this.classList.add('active');
            
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            const targetTab = document.getElementById(`${tabId}Tab`);
            if (targetTab) targetTab.classList.add('active');
            
            if (tabId === 'users') { renderizarUsuarios(); }
            if (sidebar) sidebar.classList.remove('open');
        };
    });

    // Busca
    const inputBusca = document.getElementById('searchInput');
    if (inputBusca) { inputBusca.oninput = renderizarEstoque; }

    // Importação Excel (Mantida apenas a importação, caso precise repor dados em massa)
    const btnUploadExcel = document.getElementById('uploadBtn');
    const inputArquivoExcel = document.getElementById('excelFile');
    const statusImportacao = document.getElementById('importStatus');

    if (btnUploadExcel && inputArquivoExcel) {
        btnUploadExcel.onclick = function(e) {
            e.preventDefault();
            const arquivo = inputArquivoExcel.files[0];
            if (!arquivo) { showToast('❌ Selecione uma planilha primeiro!'); return; }
            
            const leitor = new FileReader();
            leitor.onload = function(evt) {
                try {
                    const dadosBrutos = new Uint8Array(evt.target.result);
                    const livro = XLSX.read(dadosBrutos, { type: 'array' });
                    const primeiraAba = livro.Sheets[livro.SheetNames[0]];
                    const linhas = XLSX.utils.sheet_to_json(primeiraAba, { header: 1 });

                    if (linhas.length <= 1) { showToast('❌ Planilha vazia ou sem cabeçalho.'); return; }
                    let contadorSucessos = 0;

                    for (let i = 1; i < linhas.length; i++) {
                        const col = linhas[i];
                        if (!col || col.length === 0 || col[0] === "") continue; 
                        
                        estoque.push({
                            codigo: col[0] ? String(col[0]).trim() : `MAT-${i}`,       
                            nome: col[1] ? String(col[1]).trim() : 'Sem Descrição',    
                            localizacao: col[2] ? String(col[2]).trim() : 'N/D',       
                            unidade: col[3] ? String(col[3]).trim() : 'UN',            
                            quantidade: col[4] ? parseInt(col[4], 10) : 0,             
                            valor: col[5] ? parseFloat(col[5]) : 0.0,
                            categoria: col[6] ? String(col[6]).trim() : 'Geral',
                            estoqueMinimo: col[7] ? parseInt(col[7], 10) : 0
                        });
                        contadorSucessos++;
                    }

                    localStorage.setItem('estoque', JSON.stringify(estoque));
                    renderizarEstoque();
                    
                    if (statusImportacao) {
                        statusImportacao.innerHTML = `<span style="color: #10b981;">✅ Sucesso! ${contadorSucessos} itens importados.</span>`;
                    }
                    showToast('📊 Importação realizada com sucesso!');
                    inputArquivoExcel.value = '';

                } catch (erro) {
                    if (statusImportacao) statusImportacao.innerHTML = `<span style="color: #ef4444;">❌ Erro ao ler planilha.</span>`;
                }
            };
            leitor.readAsArrayBuffer(arquivo);
        };
    }

    // Cadastro de Notas 
    const formNota = document.getElementById('noteForm');
    if (formNota) {
        formNota.onsubmit = function(e) {
            e.preventDefault();
            const selectCode = document.getElementById('noteMaterialSelect').value;
            const itemOriginal = estoque.find(i => i.codigo === selectCode);
            
            if(!itemOriginal) { showToast("Erro ao localizar material no estoque."); return; }

            notasSolicitacao.unshift({
                id: Date.now().toString(),
                data: new Date().toLocaleDateString('pt-BR'),
                usuario: usuarioLogado ? usuarioLogado.nome : "Colaborador",
                itemCodigo: itemOriginal.codigo,
                itemNome: itemOriginal.nome,
                quantidade: parseInt(document.getElementById('noteQuantity').value),
                urgencia: document.getElementById('noteUrgency').value,
                status: 'Pendente' 
            });
            localStorage.setItem('notasSolicitacao', JSON.stringify(notasSolicitacao));
            renderizarNotas();
            formNota.reset();
            showToast('✅ Nota de solicitação enviada para aprovação!');
        };
    }

    // Registros e Logins
    const formRegistro = document.getElementById('registerForm');
    if (formRegistro) {
        formRegistro.onsubmit = function(e) {
            e.preventDefault();
            const userReg = document.getElementById('registerUsername').value.trim();
            const nameReg = document.getElementById('registerFullName').value.trim();
            const badgeReg = document.getElementById('registerBadge').value.trim();
            const emailReg = document.getElementById('registerEmail').value.trim();
            const compReg = document.getElementById('registerCompany').value.trim();
            const passReg = document.getElementById('registerPassword').value;

            if(userReg.toLowerCase() === 'admin' || usuariosRegistrados.some(u => u.user.toLowerCase() === userReg.toLowerCase())) {
                showToast("❌ Este nome de usuário já está cadastrado!");
                return;
            }

            usuariosRegistrados.push({
                user: userReg, nome: nameReg, cracha: badgeReg, email: emailReg, empresa: compReg, senha: passReg
            });

            localStorage.setItem('usuariosRegistrados', JSON.stringify(usuariosRegistrados));
            showToast("✅ Colaborador registrado com sucesso! Efetue o login.");
            formRegistro.reset();
            
            document.getElementById('registerScreen').classList.remove('active');
            document.getElementById('loginScreen').classList.add('active');
        };
    }

    const btnGoToRegister = document.getElementById('goToRegisterBtn');
    if (btnGoToRegister) { btnGoToRegister.onclick = (e) => { e.preventDefault(); document.getElementById('loginScreen').classList.remove('active'); document.getElementById('registerScreen').classList.add('active'); }; }

    const btnGoToLogin = document.getElementById('goToLoginBtn');
    if (btnGoToLogin) { btnGoToLogin.onclick = (e) => { e.preventDefault(); document.getElementById('registerScreen').classList.remove('active'); document.getElementById('loginScreen').classList.add('active'); }; }

    const btnLogout = document.getElementById('sidebarLogoutBtn');
    if (btnLogout) {
        btnLogout.onclick = function(e) {
            e.preventDefault();
            usuarioLogado = null;
            document.getElementById('mainScreen').classList.remove('active');
            document.getElementById('loginScreen').classList.add('active');
            document.getElementById('username').value = '';
            document.getElementById('password').value = '';
        };
    }

    const formLogin = document.getElementById('loginForm');
    if (formLogin) {
        formLogin.onsubmit = function(e) {
            e.preventDefault();
            const userIn = document.getElementById('username').value.trim();
            const passIn = document.getElementById('password').value;
            
            if (userIn.toLowerCase() === 'admin') {
                usuarioLogado = { user: "admin", nome: "Administrador Master", cracha: "0001", email: `admin@empresa.com`, empresa: "Master S.A." };
                logarNoSistema();
            } else {
                const contaEncontrada = usuariosRegistrados.find(u => u.user.toLowerCase() === userIn.toLowerCase() && u.senha === passIn);
                if (contaEncontrada) {
                    usuarioLogado = { user: contaEncontrada.user, nome: contaEncontrada.nome, cracha: contaEncontrada.cracha, email: contaEncontrada.email, empresa: contaEncontrada.empresa };
                    logarNoSistema();
                } else {
                    showToast("❌ Login inválido! Verifique seu usuário e senha.");
                }
            }
        };
    }
}

// === 4. PERMISSÕES E LOGIN ===
function logarNoSistema() {
    document.getElementById('loginScreen').classList.remove('active');
    document.getElementById('mainScreen').classList.add('active');

    document.getElementById('userDisplay').innerText = usuarioLogado.user;
    document.getElementById('profName').innerText = usuarioLogado.nome;
    document.getElementById('profCompany').innerText = usuarioLogado.empresa;
    document.getElementById('profBadge').innerText = usuarioLogado.cracha;

    if (usuarioLogado.user.toLowerCase() === 'admin') {
        document.getElementById('adminMenuOptions').style.display = 'block';
        document.getElementById('sidebarImportExportBtn').style.display = 'block';
        
        // Modal Novo Item (Apenas Admin)
        const btnAddItem = document.getElementById('sidebarAddItemBtn');
        if (btnAddItem) {
            btnAddItem.onclick = function(e) {
                e.preventDefault();
                document.getElementById('addCode').value = '';
                document.getElementById('addLocation').value = '';
                document.getElementById('addName').value = '';
                document.getElementById('addCategory').value = 'Geral';
                document.getElementById('addUnit').value = 'UN';
                document.getElementById('addQuantity').value = '';
                document.getElementById('addMinStock').value = '5';
                document.getElementById('addValue').value = '';
                document.getElementById('addItemModal').style.display = 'flex';
            };
        }

        const closeAddModal = document.getElementById('closeAddModal');
        if(closeAddModal) { closeAddModal.onclick = (e) => { e.preventDefault(); document.getElementById('addItemModal').style.display = 'none'; }; }

        const btnSaveItem = document.getElementById('saveNewItemBtn');
        if(btnSaveItem) {
            btnSaveItem.onclick = (e) => {
                e.preventDefault();
                const cod = document.getElementById('addCode').value.trim();
                const nome = document.getElementById('addName').value.trim();
                const cat = document.getElementById('addCategory').value.trim() || 'Geral';
                const loc = document.getElementById('addLocation').value.trim() || 'N/D';
                const uni = document.getElementById('addUnit').value.trim() || 'UN';
                const qtd = parseInt(document.getElementById('addQuantity').value) || 0;
                const min = parseInt(document.getElementById('addMinStock').value) || 0;
                const val = parseFloat(document.getElementById('addValue').value) || 0;

                if(!cod || !nome) { showToast('❌ Código e Descrição são obrigatórios!'); return; }
                if(estoque.find(i => i.codigo === cod)) { showToast('❌ Código já existe!'); return; }

                estoque.push({ codigo: cod, nome: nome, categoria: cat, localizacao: loc, unidade: uni, quantidade: qtd, estoqueMinimo: min, valor: val });
                localStorage.setItem('estoque', JSON.stringify(estoque));

                document.getElementById('addItemModal').style.display = 'none';
                renderizarEstoque();
                atualizarDropdownNotas();
                showToast('✅ Novo item salvo!');
            };
        }
        
        // Limpar Tudo com Validação de Senha Interna
        const btnLimparEstoque = document.getElementById('sidebarClearAllBtn');
        const passModal = document.getElementById('passwordModal');

        if (btnLimparEstoque) {
            btnLimparEstoque.onclick = function(e) {
                e.preventDefault();
                if(passModal) passModal.style.display = 'flex';
                const inp = document.getElementById('adminPasswordInput');
                if(inp) inp.focus();
            };
        }

        const btnClosePass = document.getElementById('closePassModal');
        if(btnClosePass) {
            btnClosePass.onclick = () => { if(passModal) passModal.style.display = 'none'; };
        }

        const btnConfirmPass = document.getElementById('confirmPassModal');
        if(btnConfirmPass) {
            btnConfirmPass.onclick = function() {
                const senhaInput = document.getElementById('adminPasswordInput');
                const senha = senhaInput ? senhaInput.value : '';
                
                if (senha !== "admin123") { // Altere para sua senha
                    showToast("❌ Senha incorreta!");
                    if(senhaInput) senhaInput.value = '';
                    return;
                }

                if(passModal) passModal.style.display = 'none';
                estoque = []; 
                movimentacoes = []; 
                notasSolicitacao = [];
                
                localStorage.setItem('estoque', JSON.stringify([]));
                localStorage.setItem('movimentacoes', JSON.stringify([]));
                localStorage.setItem('notasSolicitacao', JSON.stringify([]));
                
                renderizarEstoque(); 
                renderizarHistorico(); 
                renderizarNotas();
                
                showToast("🗑️ Estoque limpo com sucesso!");
                if(senhaInput) senhaInput.value = '';
            };
        }

    } else {
        // Se for um usuário comum, oculta as opções de Admin
        document.getElementById('adminMenuOptions').style.display = 'none';
        document.getElementById('sidebarImportExportBtn').style.display = 'none';
    }

    document.querySelectorAll('.sidebar-link').forEach(l => l.classList.remove('active'));
    const linkInventario = document.querySelector('[data-tab="inventory"]');
    if(linkInventario) linkInventario.classList.add('active');
    
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    const tabInventario = document.getElementById('inventoryTab');
    if(tabInventario) tabInventario.classList.add('active');

    atualizarDropdownNotas();
    renderizarEstoque();
    renderizarHistorico();
    renderizarNotas();
}

// === 5. RENDERIZAÇÃO INTELIGENTE DE ESTOQUE ===
function renderizarEstoque() {
    const listContainer = document.getElementById('itemsList');
    const inputBusca = document.getElementById('searchInput');
    const filtroTexto = inputBusca ? inputBusca.value.toLowerCase() : '';
    
    if (!listContainer) return;
    listContainer.innerHTML = '';
    
    // Filtro Combinado (Busca em Texto + Categoria Clicada)
    let itensFiltrados = estoque.filter(item => {
        const textoMatch = item.codigo.toLowerCase().includes(filtroTexto) || item.nome.toLowerCase().includes(filtroTexto) || item.localizacao.toLowerCase().includes(filtroTexto);
        const catMatch = (categoriaAtual === 'Todas' || item.categoria === categoriaAtual);
        return textoMatch && catMatch;
    });

    const emptyMessage = document.getElementById('emptyMessage');
    if(emptyMessage) emptyMessage.style.display = itensFiltrados.length === 0 ? 'block' : 'none';

    let totalQtdFisica = 0;
    let alertasCount = 0;

    itensFiltrados.forEach(item => {
        totalQtdFisica += item.quantidade;
        
        // Verifica Estoque Mínimo
        const isLowStock = item.quantidade <= (item.estoqueMinimo || 0);
        if(isLowStock) alertasCount++;

        const card = document.createElement('div');
        card.className = `item-card ${isLowStock ? 'low-stock' : ''}`;
        
        // Monta o Card
        card.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:10px;">
                <h4 style="margin:0; font-size:1rem; flex:1;">${item.nome} <br><span style="font-size:0.7rem; color:var(--text-light); font-weight:normal;">${item.categoria || 'Geral'}</span></h4>
                <span class="item-badge">${item.localizacao}</span>
            </div>
            ${isLowStock ? `<div style="background:#fef2f2; border:1px solid #fecaca; color:#ef4444; padding:4px 8px; font-size:0.75rem; border-radius:4px; font-weight:bold; margin-bottom:8px; display:inline-block;">⚠️ ESTOQUE BAIXO</div>` : ''}
            <p style="font-size:0.85rem; margin:3px 0;">Código: <strong>${item.codigo}</strong></p>
            <p style="font-size:0.85rem; margin:3px 0;">Saldo: <strong style="color: ${isLowStock ? '#ef4444' : 'var(--primary-color)'};">${item.quantidade} ${item.unidade}</strong> ${item.estoqueMinimo ? `<small style="color:var(--text-light)">(Mín: ${item.estoqueMinimo})</small>` : ''}</p>
            <p style="font-size:0.85rem; margin:3px 0;">Valor: <strong>R$ ${item.valor.toFixed(2)}</strong></p>
        `;
        card.onclick = () => abrirModalMovimentacao(item);
        listContainer.appendChild(card);
    });

    // Atualiza Painel Numérico
    const totalItems = document.getElementById('totalItems');
    if(totalItems) totalItems.innerText = totalQtdFisica;
    
    const stockedItems = document.getElementById('stockedItems');
    if(stockedItems) stockedItems.innerText = itensFiltrados.length;
    
    const lowStockAlerts = document.getElementById('lowStockAlerts');
    if(lowStockAlerts) lowStockAlerts.innerText = alertasCount;

    // Renderiza Botões de Categoria
    renderizarFiltrosCategoria();
}

function renderizarFiltrosCategoria() {
    const container = document.getElementById('categoryFilters');
    if(!container) return;
    
    // Mapeia todas as categorias existentes
    const categoriasExistem = [...new Set(estoque.map(i => i.categoria || 'Geral'))];
    categoriasExistem.unshift('Todas');

    container.innerHTML = '';
    categoriasExistem.forEach(cat => {
        const btn = document.createElement('button');
        btn.className = `cat-btn ${cat === categoriaAtual ? 'active' : ''}`;
        btn.innerText = cat;
        btn.onclick = () => { categoriaAtual = cat; renderizarEstoque(); };
        container.appendChild(btn);
    });
}

function atualizarDropdownNotas() {
    const sel = document.getElementById('noteMaterialSelect');
    if(!sel) return;
    sel.innerHTML = '<option value="">-- Selecione o Material --</option>';
    estoque.forEach(i => {
        const opt = document.createElement('option');
        opt.value = i.codigo;
        opt.innerText = `${i.codigo} - ${i.nome} (Disp: ${i.quantidade})`;
        sel.appendChild(opt);
    });
}

function renderizarHistorico() {
    const container = document.getElementById('historyList');
    const adminSearch = document.getElementById('adminHistorySearch');
    const adminInput = document.getElementById('adminHistoryInput');
    
    if (!container || !usuarioLogado) return;
    
    container.innerHTML = '';
    const filtro = adminInput ? adminInput.value.toLowerCase() : '';
    const eAdmin = usuarioLogado.user.toLowerCase() === 'admin';

    // Lógica: Filtra para mostrar só o do usuário OU tudo para o admin
    let listaFiltrada = movimentacoes.filter(mov => {
        if (!eAdmin) return mov.colaborador.toLowerCase() === usuarioLogado.nome.toLowerCase();
        return mov.colaborador.toLowerCase().includes(filtro) || 
               mov.item.toLowerCase().includes(filtro);
    });

    if (adminSearch) adminSearch.style.display = eAdmin ? 'block' : 'none';
    
    const emptyMsg = document.getElementById('emptyHistory');
    if (emptyMsg) emptyMsg.style.display = listaFiltrada.length === 0 ? 'block' : 'none';

    listaFiltrada.forEach(mov => {
        const div = document.createElement('div');
        div.className = 'history-item';
        div.style.borderLeft = `4px solid ${mov.tipo === 'SAÍDA' ? '#ef4444' : '#10b981'}`;
        div.innerHTML = `<strong>${mov.tipo}</strong> - ${mov.item} (Qtd: ${mov.qtd})<br>
                         <small>${mov.colaborador} | ${mov.data}</small><br>
                         <small style="color:var(--text-light);">Motivo: ${mov.motivo}</small>`;
        container.appendChild(div);
    });
}

// === 6. PAINEL DE NOTAS (COM SISTEMA DE APROVAÇÃO) ===
function renderizarNotas() {
    const container = document.getElementById('notesDisplayList');
    const badge = document.getElementById('badgeNotasPendente');
    if(!container) return;
    
    container.innerHTML = '';
    const emptyNotes = document.getElementById('emptyNotesMessage');
    if(emptyNotes) emptyNotes.style.display = notasSolicitacao.length === 0 ? 'block' : 'none';

    let pendentesCount = 0;

    notasSolicitacao.forEach(nota => {
        if(nota.status === 'Pendente') pendentesCount++;

        const div = document.createElement('div');
        div.className = 'note-card';
        
        let corBadge = '#3b82f6';
        if(nota.urgencia === 'Média') corBadge = '#eab308';
        if(nota.urgencia === 'Alta') corBadge = '#ef4444';

        let corStatus = nota.status === 'Aprovada' ? '#10b981' : (nota.status === 'Recusada' ? '#ef4444' : '#f59e0b');

        // Layout do card de nota
        div.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom: 5px;">
                <div>
                    <strong style="font-size:0.95rem;">${nota.itemNome} (x${nota.quantidade})</strong><br>
                    <small>Cód: ${nota.itemCodigo}</small>
                </div>
                <div style="display:flex; flex-direction:column; gap:5px; align-items:flex-end;">
                    <span style="font-size:0.7rem; background:${corBadge}; color:white; padding:2px 6px; border-radius:4px; font-weight:bold;">${nota.urgencia}</span>
                    <span style="font-size:0.7rem; color:${corStatus}; font-weight:bold;">${nota.status}</span>
                </div>
            </div>
            <small style="color:var(--text-light); display:block; margin-top:10px;">Solicitante: <strong>${nota.usuario}</strong> - ${nota.data}</small>
        `;

        // Se for admin e estiver pendente, mostra os botões de aprovação
        if(usuarioLogado && usuarioLogado.user.toLowerCase() === 'admin' && nota.status === 'Pendente') {
            const btnGroup = document.createElement('div');
            btnGroup.style.display = 'flex'; btnGroup.style.gap = '10px'; btnGroup.style.marginTop = '15px';
            
            const btnAprovar = document.createElement('button');
            btnAprovar.innerText = '✅ Aprovar'; btnAprovar.className = 'btn'; btnAprovar.style.background = '#10b981'; btnAprovar.style.color = 'white'; btnAprovar.style.flex = '1';
            btnAprovar.onclick = () => aprovarNota(nota.id);

            const btnRecusar = document.createElement('button');
            btnRecusar.innerText = '❌ Recusar'; btnRecusar.className = 'btn'; btnRecusar.style.background = '#ef4444'; btnRecusar.style.color = 'white'; btnRecusar.style.flex = '1';
            btnRecusar.onclick = () => recusarNota(nota.id);

            btnGroup.appendChild(btnAprovar); btnGroup.appendChild(btnRecusar);
            div.appendChild(btnGroup);
        }
        
        container.appendChild(div);
    });

    if(badge) {
        badge.innerText = pendentesCount;
        badge.style.display = pendentesCount > 0 ? 'inline-block' : 'none';
    }
}

function aprovarNota(notaId) {
    const nota = notasSolicitacao.find(n => n.id === notaId);
    if(!nota) return;

    const itemEst = estoque.find(i => i.codigo === nota.itemCodigo);
    if(!itemEst) { showToast("O material não existe mais no estoque."); return; }
    if(itemEst.quantidade < nota.quantidade) { showToast(`Saldo insuficiente! Há apenas ${itemEst.quantidade} UN disponíveis.`); return; }

    // Deduz do estoque
    itemEst.quantidade -= nota.quantidade;
    nota.status = 'Aprovada';

    // Registra Movimento
    movimentacoes.unshift({
        data: new Date().toLocaleString('pt-BR'),
        colaborador: nota.usuario,
        cracha: "Via App",
        tipo: 'SAÍDA', item: itemEst.nome, qtd: nota.quantidade, motivo: "Nota de Solicitação Aprovada"
    });

    localStorage.setItem('estoque', JSON.stringify(estoque));
    localStorage.setItem('notasSolicitacao', JSON.stringify(notasSolicitacao));
    localStorage.setItem('movimentacoes', JSON.stringify(movimentacoes));

    renderizarEstoque(); renderizarHistorico(); renderizarNotas(); atualizarDropdownNotas();
}

function recusarNota(notaId) {
    const nota = notasSolicitacao.find(n => n.id === notaId);
    if(nota) {
        nota.status = 'Recusada';
        localStorage.setItem('notasSolicitacao', JSON.stringify(notasSolicitacao));
        renderizarNotas();
    }
}

// === 7. GERENCIAMENTO DE USUÁRIOS (SÓ ADMIN) ===
function renderizarUsuarios() {
    const cont = document.getElementById('usersListContainer');
    if(!cont) return;
    cont.innerHTML = '';

    if(usuariosRegistrados.length === 0) {
        cont.innerHTML = '<div class="empty-state">Nenhum colaborador comum cadastrado.</div>'; return;
    }

    usuariosRegistrados.forEach((u, index) => {
        const div = document.createElement('div');
        div.style.padding = '15px'; div.style.border = '1px solid var(--border-color)'; div.style.borderRadius = '8px'; div.style.display = 'flex'; div.style.justifyContent = 'space-between'; div.style.alignItems = 'center';
        
        div.innerHTML = `
            <div>
                <strong style="font-size:1.05rem;">${u.nome}</strong> (Crachá: ${u.cracha})<br>
                <small style="color:var(--text-light);">Login: <strong>${u.user}</strong> | Setor: ${u.empresa}</small>
            </div>
        `;

        const delBtn = document.createElement('button');
        delBtn.innerText = 'Excluir Conta';
        delBtn.className = 'btn btn-danger';
        delBtn.onclick = () => {
            if(confirm(`Deseja realmente remover o acesso de ${u.nome}?`)) {
                usuariosRegistrados.splice(index, 1);
                localStorage.setItem('usuariosRegistrados', JSON.stringify(usuariosRegistrados));
                renderizarUsuarios();
            }
        };

        div.appendChild(delBtn);
        cont.appendChild(div);
    });
}

// === 8. MODAL DE MOVIMENTAÇÃO MANUAL ===
function abrirModalMovimentacao(item) {
    itemSelecionadoParaMover = item;
    const modalTitle = document.getElementById('modalTitle');
    if(modalTitle) modalTitle.innerText = `Movimentar: ${item.nome}`;
    
    document.getElementById('modalCode').value = item.codigo;
    document.getElementById('modalLocation').value = item.localizacao;
    document.getElementById('modalName').value = item.nome;
    document.getElementById('modalUnit').value = item.unidade;
    document.getElementById('modalQuantity').value = item.quantidade;
    document.getElementById('modalValue').value = `R$ ${item.valor.toFixed(2)}`;
    document.getElementById('moveQuantity').value = '';
    document.getElementById('moveReason').value = '';
    
    const itemModal = document.getElementById('itemModal');
    if(itemModal) itemModal.style.display = 'flex';
}

const btnCloseModal = document.getElementById('closeModal');
if(btnCloseModal) btnCloseModal.onclick = function(e) { e.preventDefault(); document.getElementById('itemModal').style.display = 'none'; };

const btnExit = document.getElementById('exitBtn');
if(btnExit) btnExit.onclick = function(e) { e.preventDefault(); processarMovimento('SAÍDA'); };

const btnEntry = document.getElementById('entryBtn');
if(btnEntry) btnEntry.onclick = function(e) { e.preventDefault(); processarMovimento('ENTRADA'); };

function processarMovimento(tipo) {
    const qtdInput = parseInt(document.getElementById('moveQuantity').value);
    const motivoInput = document.getElementById('moveReason').value.trim() || "Manual";

    if (!qtdInput || qtdInput <= 0) { showToast('❌ Quantidade inválida!'); return; }
    if (tipo === 'SAÍDA' && qtdInput > itemSelecionadoParaMover.quantidade) { showToast('❌ Saldo insuficiente!'); return; }

    itemSelecionadoParaMover.quantidade += (tipo === 'ENTRADA' ? qtdInput : -qtdInput);

    movimentacoes.unshift({
        data: new Date().toLocaleString('pt-BR'),
        colaborador: usuarioLogado ? usuarioLogado.nome : "Usuário",
        cracha: usuarioLogado ? usuarioLogado.cracha : "0000",
        tipo: tipo, item: itemSelecionadoParaMover.nome, qtd: qtdInput, motivo: motivoInput
    });

    localStorage.setItem('estoque', JSON.stringify(estoque));
    localStorage.setItem('movimentacoes', JSON.stringify(movimentacoes));
    
    const itemModal = document.getElementById('itemModal');
    if(itemModal) itemModal.style.display = 'none';
    
    renderizarEstoque(); 
    renderizarHistorico(); 
    atualizarDropdownNotas();
}

// START VAI!
iniciarSistema();
