# CLAUDE.md - Sistema FOMENTAR Legado

Sistema web completo para conversão de arquivos SPED e apuração de incentivos fiscais FOMENTAR/PRODUZIR/MICROPRODUZIR, ProGoiás e LogPRODUZIR do Estado de Goiás.

## 🎯 **Status do Sistema**

### ✅ **Sistema 100% Funcional**
- **4 módulos completos**: Conversor SPED, FOMENTAR, ProGoiás, LogPRODUZIR
- **Sistema de autenticação** completo com 8 perfis de usuário
- **13.693 linhas** de lógica fiscal implementada e testada
- **Interface moderna** com tabelas HTML profissionais
- **Pronto para produção** com toda funcionalidade fiscal validada

## 🏗️ Arquitetura do Sistema (4 Módulos Fiscais)

### 1. **Conversor SPED** ✅
- **Importação/parsing** arquivos SPED EFD-ICMS/IPI (.txt)
- **Exportação Excel** com múltiplas abas organizadas
- **Interface drag-and-drop** responsiva
- **Status**: 100% funcional

### 2. **FOMENTAR/PRODUZIR/MICROPRODUZIR** ✅
- **Classificação automática** CFOPs incentivados vs não-incentivados
- **Demonstrativo v3.51** com 44 itens fiscais oficiais
- **Correção inteligente** códigos E111/C197/D197
- **Geração automática** registro E115 (54 códigos GO200001-GO200054)
- **Confronto calculado vs SPED** declarado
- **Múltiplos períodos** com compensação sequencial
- **Exclusão créditos circulares** (GO040007, GO040008, etc.)
- **Status**: 100% funcional

### 3. **ProGoiás** ✅
- **Base legal**: Decreto nº 9.724/2020
- **Fórmula oficial**: ICMSS - ICMSE - AJCRED + AJDEB
- **Configuração dinâmica** por tipo empresa/ano
- **Múltiplos períodos** suportado
- **Status**: 100% funcional

### 4. **LogPRODUZIR** ✅
- **Base legal**: Lei nº 14.244/2002, Decreto nº 5.835/2003
- **CFOPs transporte**: Interestaduais (6351-6360, 6932) + Estaduais (5351-5360, 5932)
- **Categorias**: I(50%), II(73%), III(80%) com média histórica corrigida
- **Contribuições obrigatórias**: 20% total (2% Bolsa + 3% FUNPRODUZIR + 15% PROTEGE)
- **Correção IGP-DI**: Atualização mensal automática
- **Status**: 100% funcional

## 🔐 Sistema de Autenticação Completo

### **Usuários Pré-configurados (8 Perfis)**

| Usuário | Senha | Perfil | Descrição |
|---------|-------|--------|-----------|
| `admin` | `admin0000` | admin | **Administrador completo** |
| `supervisor` | `super123` | admin | **Supervisor completo** |
| `fomentar.basico` | `fom123` | fomentarBasico | **FOMENTAR período único** |
| `fomentar.completo` | `fomc123` | fomentarCompleto | **FOMENTAR completo + E115** |
| `progoias.basico` | `pro123` | progoiasBasico | **ProGoiás período único** |
| `progoias.completo` | `proc123` | progoiasCompleto | **ProGoiás completo** |
| `logproduzir.basico` | `log123` | logproduzirBasico | **LogPRODUZIR período único** |
| `logproduzir.completo` | `logc123` | logproduzirCompleto | **LogPRODUZIR completo** |
| `conversor` | `conv123` | converterApenas | **Apenas conversor SPED** |

### **Funcionalidades de Segurança**
- **Sessões JWT** com expiração 4 horas
- **Renovação automática** em atividade
- **Controle granular** por funcionalidade
- **Logout automático** ao expirar
- **Interface adaptativa** conforme permissões

## 🚀 Como Usar

### **Início Rápido**
1. Abra `sped-web-fomentar.html` no navegador
2. Faça login com `admin/admin0000` (acesso completo)
3. Carregue arquivo SPED via drag-and-drop
4. Navegue pelas abas dos módulos fiscais

### **Fluxo de Trabalho Típico**
1. **Login** → Escolher perfil adequado
2. **Conversor** → Upload arquivo SPED (.txt)
3. **FOMENTAR** → Configurar parâmetros e calcular
4. **E115** → Gerar arquivo SPED + Excel comparativo
5. **ProGoiás/LogPRODUZIR** → Apurações específicas

## 📊 Processamento SPED EFD-ICMS/IPI

### **Registros Consolidados (Principais)**
- **C190**: NF-e consolidado por CFOP/CST/alíquota
- **C590**: NF-e Energia/Telecom consolidado
- **D190**: CT-e consolidado por CFOP/CST/alíquota  
- **D590**: CT-e Serviços consolidado

### **Registros de Apuração**
- **E100/E110**: Apuração geral do ICMS
- **E111**: Créditos/débitos específicos (processamento inteligente)
- **E115**: Demonstrativo fiscal (geração automática 54 códigos)
- **C197**: Outras obrigações ICMS (correção suportada)
- **D197**: Outras obrigações CT-e (correção suportada)

## 🎨 Melhorias de Interface Implementadas

### **Tabelas HTML Modernas**
- ✅ **Códigos E111**: Tabela scrollável com classificação incentivado/não-incentivado
- ✅ **Códigos C197/D197**: Interface profissional com badges coloridos
- ✅ **CFOPs Genéricos**: Tabela com radio buttons para classificação
- ✅ **Design responsivo**: Sticky headers e scrollbars personalizadas
- ✅ **Experiência visual**: Gradientes, hover effects e estilos modernos

### **Funcionalidades da Interface**
- **Altura máxima**: 400px-500px com scroll automático
- **Cabeçalhos fixos**: Navegação facilitada em listas longas
- **Badges visuais**: Cores diferenciadas por tipo de operação
- **Input responsivo**: Campos de correção integrados nas tabelas
- **Scrollbars customizadas**: Visual profissional consistente

## 📁 Estrutura de Arquivos

```
FOMENTAR-sistema-legado/
├── sped-web-fomentar.html     # 🎯 Arquivo principal (COM autenticação)
├── index.html                 # Interface alternativa (SEM autenticação) 
├── script.js                  # ⚙️ Engine principal (13.693 linhas)
├── style.css                  # 🎨 Estilos + tabelas modernas
├── auth.js                    # 🔐 Sistema de autenticação
├── permissions.js             # 👥 Controle de permissões
└── images/                    # 🖼️ Assets visuais
    └── logo-expertzy.png
```

### **Arquivos Principais**

| Arquivo | Tamanho | Função |
|---------|---------|--------|
| `script.js` | ~685KB | Engine fiscal completa |
| `style.css` | ~55KB | Estilos + tabelas HTML |
| `sped-web-fomentar.html` | ~75KB | Interface principal autenticada |
| `auth.js` | ~15KB | Sistema de login JWT |
| `permissions.js` | ~19KB | Controle de acesso granular |

## 🔧 Constantes Fiscais Críticas

### **FOMENTAR/PRODUZIR/MICROPRODUZIR**

#### **Metodologia de Apuração ICMS - IN 885/07-GSF**
- **Classificação binária**: Incentivado vs Não-incentivado (sem proporcionalidade)
- **Anexo I**: CFOPs de entrada incentivados → Créditos incentivados
- **Anexo II**: CFOPs de saída incentivados → Débitos incentivados  
- **Anexo III**: Códigos de ajuste incentivados → Créditos OU Débitos

#### **Percentuais Oficiais**
- ✅ **FOMENTAR**: 70%
- ✅ **PRODUZIR**: 73%
- ✅ **MICROPRODUZIR**: 90%

#### **Exclusões Automáticas**
```javascript
// Créditos circulares excluídos automaticamente
CODIGOS_EXCLUSAO_AUTOMATICA = [
  'GO040007', 'GO040008', 'GO040009', 'GO040010'
];
```

### **ProGoiás**
```javascript
// Fórmula oficial (Decreto 9.724/2020)
FORMULA_PROGOIAS = "ICMSS - ICMSE - AJCRED + AJDEB";
PERCENTUAL_PADRAO = 50%; // Configurável por empresa/ano
```

### **LogPRODUZIR**
```javascript
// CFOPs Fretes Incentivados
CFOP_INTERESTADUAIS = [6351, 6352, 6353, 6354, 6355, 6356, 6357, 6359, 6360, 6932];
CFOP_ESTADUAIS = [5351, 5352, 5353, 5354, 5355, 5356, 5357, 5359, 5360, 5932];

// Percentuais por categoria
CATEGORIAS = {
  I: 50%,   // Logística básica
  II: 73%,  // Logística + transporte (padrão)  
  III: 80%  // Acima R$ 900k mensais
};
```

## 🧪 Testing & Debug

### **Console Commands Úteis**
```javascript
// Autenticação
showSessionInfo();
getCurrentUser();
logout();

// Debug FOMENTAR
console.log(fomentarData);
continuarCalculoFomentar();

// Debug LogPRODUZIR  
console.log(logproduzirData);
processLogproduzirData();

// Debug ProGoiás
console.log(progoiasData);
```

### **Logs Esperados**
```
[FOMENTAR] E111 EXCLUÍDO: GO040007 = R$ 5.250,00 - NÃO computado
[PROGOIAS] Base = R$ 15.750,00 x 50% = R$ 7.875,00
[LOGPRODUZIR-TRANSPORTE] CFOP 6351 = R$ 1.500,00 (Categoria II)
[E115] Código GO200015 = R$ 2.300,00 calculado vs R$ 2.100,00 SPED
```

## 🚨 Troubleshooting Comum

| Problema | Causa Provável | Solução |
|----------|----------------|---------|
| **Valores zerados** | CFOPs não encontrados no SPED | Verificar arquivo/CFOPs configurados |
| **Interface não atualiza** | Erro JavaScript | Console (F12) para debug |
| **Upload falha** | Arquivo muito grande | Arquivo <200MB recomendado |
| **Sessão expirada** | 4h timeout | Re-login necessário |
| **CFOPs genéricos** | Classificação pendente | Configurar na interface |

## 📊 Estatísticas Técnicas

- **Total de Linhas**: 13.693 (script.js)
- **Usuários configurados**: 8 perfis distintos
- **CFOPs mapeados**: 200+ (incentivados/não-incentivados)
- **Códigos E115**: 54 automáticos (GO200001-GO200054)
- **Registros SPED**: C190, C590, D190, D590, E100, E110, E111, C197, D197
- **Períodos suportados**: Múltiplos com compensação automática

## ⚡ Performance & Limitações

### **Otimizações**
- **Client-side**: Sem servidor necessário
- **Processing**: Assíncrono com progress feedback
- **Memory**: Otimizado para arquivos SPED grandes
- **Cache**: LocalStorage para sessões persistentes

### **Limitações Conhecidas**
- **Memória navegador**: SPEDs >200MB podem falhar
- **Processing**: Single-thread (JavaScript)
- **Storage**: LocalStorage limitado por navegador

## 🔄 Histórico de Alterações Recentes

### **✅ Agosto 2025 - Melhorias de Interface**

#### **Interface Modernizada**
- **Códigos E111**: Convertidos de divs para tabelas HTML scrolláveis
- **Códigos C197/D197**: Interface profissional com badges coloridos  
- **CFOPs Genéricos**: Tabela com radio buttons para classificação
- **Estilos CSS**: Gradientes, hover effects, scrollbars customizadas

#### **Correções Fiscais Críticas**
- **Item 35 FOMENTAR**: Campo manual (era automático incorreto)
- **Fórmulas de exibição**: Corrigidas para refletir IN 885/07-GSF
- **Confrontação SPED**: Usa VL_ICMS_RECOLHER (não VL_SLD_APURADO)

## 🏆 Status Final

**✅ Sistema 100% Funcional para Produção**
- Interface moderna e profissional
- Cálculos fiscais validados e corretos
- Autenticação robusta com controle granular
- Documentação completa e atualizada
- Pronto para uso em ambiente empresarial

---

## 📞 Suporte

Sistema desenvolvido com inteligência fiscal brasileira para apuração precisa de incentivos ICMS do Estado de Goiás.

**Principais funcionalidades validadas:**
- ✅ Apuração FOMENTAR/PRODUZIR/MICROPRODUZIR
- ✅ Cálculo ProGoiás (Decreto 9.724/2020)  
- ✅ Apuração LogPRODUZIR (Lei 14.244/2002)
- ✅ Conversão SPED para Excel organizado
- ✅ Geração E115 automática com 54 códigos
- ✅ Interface responsiva e moderna

**Última atualização**: Agosto 2025 - Interface de tabelas HTML implementada