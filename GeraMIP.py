# ============================================================================================
# Modulo de Geração da Matriz Insumo Produto Nacional ESTIMADA  a partir das TABELAS TRUs do IBGE
# Desenvolvido por João Maria de Oliveira
# Baseado em Guilhoto (2010), IBGE (2010) e Barry-Miller (2009)
# Conforme o sistema de contas Nacionais 3a edição
# ============================================================================================
import numpy as np
import sys
import FuncoesApoio as funcaoApoio

# ============================================================================================
# parâmetros para a construção da MIP
# ============================================================================================
# nDimensao - tamanho da diensão da matriz Insumo produto
#   Valores possíveis:  0 - 12x12; 1 - 20x20; 2 - 107X51; 3 - 128x68
# vRowsTrade  - Números das linhas (inicial e final), dos produtos relacionados ao comercio
# vRowTransp - Números das linhas (inicial e final), dos produtos reacionados ao  tranporte
# vColsTrade  - Números das colunas (inicial e final), das atividades reacionadas ao comercio
# vColsTransp - Números das  colunas (inicial e final), das atividades reacionadas ao  tranporte
#  - a posição em cada vetor é dada por nDimensao ( Tamanho da MIP)
# nProduct - Número de produtos de acordo com a dimensão da MIP
# nSector - Número de atividades de acordo com a dimensão da MIP
# lAdjustMargins - True se ajusta as margens de comércio e transporte para apenas um produto e uma atividade

vProducts = [12, 20, 107, 128]
vSectors  = [12, 20,  51,  68]
vRowsTrade  = [[5, 5], [6, 6], [88, 88], [92, 93]]
vRowsTransp = [[6, 6], [7, 7], [89, 90], [94, 97]]
vColsTrade  = [[5, 5], [6, 6], [36, 36], [40, 41]]
vColsTransp = [[6, 6], [7, 7], [37, 37], [42, 44]]
nDimensao = 3
nAno = 2015
nProduct = vProducts [nDimensao]
nSector = vSectors [nDimensao]
vRowsTradeElim  = vRowsTrade[nDimensao]
vRowsTranspElim = vRowsTransp[nDimensao]
vColsTradeElim  = vColsTrade[nDimensao]
vColsTranspElim = vColsTransp[nDimensao]

#lAdjustMargins = True
lAdjustMargins = False
if lAdjustMargins:
   sAdjustMargins = '_Agreg'
else:
    sAdjustMargins = ''

# nColsDemand - Número de Colunas da demanda na tabela demanda do IBGE
# nColsOffer - Número de Colunas da oferta na tabela de oferta do IBGE
# nRowsAV - Número de linhas do VA na tabela VA do IBGE
# nRowTotalProduction - Número da linha do total da produção
# nColTotalDemand - Numero da coluna da demanda total na  tabela demanda
# nColFinalDemand - Numero da coluna da demanda total na  tabela demanda
# nColExport - Numero da coluna de exportação na tabela oferta
# nColFBCF - Numero da coluna de FCBF na tabela oferta
# nColEstockVar - Numero da coluna de variação de estoque na tabela oferta
# nColMarginTrade = Numero da coluna da margem de comércio na tabela oferta
# nColMarginTransport = Numero da coluna da margem de transporte na tabela oferta
# nColIPI = Numero da coluna do IPI na tabela oferta
# nColICMS = Numero da coluna do ICMS na tabela oferta
# nColOtherTaxes = Numero da coluna dos outros impostos  na tabela oferta
# nColImport = Numero da coluna dos dados de importacao na tabela oferta
# nColImportTax = Numero da coluna dos impostos de importacao na tabela oferta

nColsDemand = 8
nColsOffer = 9
nRowsAV = 14
nRowTotalProduction = nRowsAV - 2
nColTotalDemand = nColsDemand - 1
nColFinalDemand = nColsDemand - 2
nColExport = 0
nColFBCF = 4
nColEstockVar = 5
nColMarginTrade = 1
nColMarginTransport = 2
nColIPI = 4
nColICMS = 5
nColOtherTaxes = 6
nColImport = 0
nColImportTax = 3

# ============================================================================================
# parâmetros gerais
# ============================================================================================
# sDirectoryInput - Pasta de entrada de dados
# sDirectoryOutput - Pasta de Saída de dados
# sFileUses             - Arquivo de usos - Demanda
# sSheetIntermedConsum  - Planilha de consumo intermediário
# sSheetDemand          - Planilha da demanda
# sSheetValueAdded      - Planilha do Valor adicionado
# sFileResources        - Arquivo de recursos
# sSheetOffer           - Planilha da oferta
# sSheetProduction      - Planilha da produção
# sSheetImport          - Planilha de importação

# sFileSheet - nome do arquivo de saida contendo as = tabelas
# vDataSheet  - Vetor contendo os  dados das planilhas
# vSheetName - Vetor contendo os nomes das planilhas
# vRowsLabel - Vetor contendo os titulos das linhas
# vColsLabel = Vetor contendo os titulos das colunas

sDirectoryInput  = './Input/'
sDirectoryOutput = './Output/'
sFileUses               = str(nSector)+'_tab2_'+str(nAno)+'.xls'
sSheetIntermedConsum    = 'CI'
sSheetDemand            = 'demanda'
sSheetAddedValue        = 'VA'
sFileResources          = str(nSector)+'_tab1_'+str(nAno)+'.xls'
sSheetOffer             = 'oferta'
sSheetProduction        = 'producao'
sSheetImport            = 'importacao'
sFileSheet = 'MIP_'+str(nAno)+'_'+str(nSector)+sAdjustMargins+'.xlsx'
vDataSheet = []
vSheetName = []
vRowsLabel = []
vColsLabel = []


# ============================================================================================
# Import values from TRUs
# ============================================================================================
vCodProduct, vNameProduct, vCodSector, vNameSector, mIntermConsum = funcaoApoio.load_intermediate_consumption\
                (sDirectoryInput, sFileUses, sSheetIntermedConsum, nProduct, nSector)

mDemand, vNameDemand = funcaoApoio.load_demand(sDirectoryInput, sFileUses, sSheetDemand, nProduct, nColsDemand)

mAddedValue, vNameAddedValue = funcaoApoio.load_gross_added_value\
                (sDirectoryInput, sFileUses, sSheetAddedValue, nSector, nRowsAV)

mOffer, vNameOffer = funcaoApoio.load_offer(sDirectoryInput, sFileResources, sSheetOffer, nProduct, nColsOffer)

mProduction = funcaoApoio.load_production(sDirectoryInput, sFileResources, sSheetProduction, nProduct, nSector)

vImport = funcaoApoio.load_import(sDirectoryInput, sFileResources, sSheetImport, nProduct)

# ============================================================================================
# Adjusting Trade and Transport for Products and for Sectors
#  ============================================================================================
lAdjust = lAdjustMargins
nAdjust = 0
while lAdjust:
    if nAdjust ==0:
        nRowIni=vRowsTradeElim[0]
        nRowFim=vRowsTradeElim[1]
        vNameProduct[nRowIni] = 'Comércio'
        nColIni = vColsTradeElim[0]
        nColFim = vColsTradeElim[1]
        vNameSector[nColIni] = 'Comércio'
    else:
        nRowIni = vRowsTranspElim[0]
        nRowFim = vRowsTranspElim[1]
        vNameProduct[nRowIni] = 'Transporte'
        nColIni = vColsTranspElim[0]
        nColFim = vColsTranspElim[1]
        vNameSector[nColIni] = 'Transporte'

    for nElim in range(nRowIni+1, nRowFim + 1):
        vNameProduct[nElim] = 'x'

    for nElim in range(nRowIni+1,nRowFim+1):
        vImport[nRowIni] += vImport[nElim]
        vImport[nElim] = 0.0

    for i in range(nColsOffer):
        for nElim in range(nRowIni+1, nRowFim + 1):
            mOffer[nRowIni, i] += mOffer[nElim, i]
            mOffer[nElim, i] = 0.0

    for i in range(nSector+1):
        for nElim in range(nRowIni+1, nRowFim + 1):
            mProduction[nRowIni, i] += mProduction[nElim, i]
            mProduction[nElim, i] = 0.0
            mIntermConsum[nRowIni, i] += mIntermConsum[nElim, i]
            mIntermConsum[nElim, i] = 0.0

    for i in range(nColsDemand):
          for nElim in range(nRowIni+1, nRowFim + 1):
              mDemand[nRowIni, i] += mDemand[nElim, i]
              mDemand[nElim, i] = 0.0


    for nElim in range(nColIni+1, nColFim + 1):
        vNameSector[nElim] = 'x'

    for i in range(nRowsAV):
        for nElim in range(nColIni+1,nColFim+1):
            mAddedValue[i, nColIni] += mAddedValue[i, nElim]
            mAddedValue[i, nElim] = 0.0

    for i in range(nProduct+1):
        for nElim in range(nColIni+1,nColFim+1):
            mProduction[i, nColIni] += mProduction[i, nElim]
            mProduction[i, nElim] = 0.0
            mIntermConsum[i, nColIni] += mIntermConsum[i, nElim]
            mIntermConsum[i, nElim] = 0.0

    nAdjust +=1
    if nAdjust == 2:
        lAdjust = False

# ============================================================================================
# Calculanting Coeficients without Stock Variation
#  ============================================================================================
mDemandWithoutEstock = np.copy(mDemand)
# zerando coluna de variação de estoque
mDemandWithoutEstock[:, nColEstockVar] = 0.0
# Diminuindo as exportações da Demanda Final e da Demanda Total
for p in range(nProduct+1):
    mDemandWithoutEstock[p, nColTotalDemand] = mDemand[p, nColTotalDemand] - mDemand[p, nColEstockVar]
    mDemandWithoutEstock[p, nColFinalDemand] = mDemand[p, nColFinalDemand] - mDemand[p, nColEstockVar]

mDistribution, mTotalConsum = funcaoApoio.calculation_distribution_matrix(mIntermConsum, mDemandWithoutEstock)

# ============================================================================================
# Calculanting Arrays internally distributed by alphas
#  ============================================================================================
nColMarginTrade = 1
mMarginTrade = funcaoApoio.calculation_margin(mDistribution, mOffer, nColMarginTrade, vRowsTradeElim)

nColMarginTransport = 2
mMarginTransport = funcaoApoio.calculation_margin(mDistribution, mOffer, nColMarginTransport, vRowsTranspElim)

nColIPI = 4
mIPI = funcaoApoio.calculation_internal_matrix(mDistribution, mOffer, nColIPI)

nColICMS = 5
mICMS = funcaoApoio.calculation_internal_matrix(mDistribution, mOffer, nColICMS)

nColOtherTaxes = 6
mOtherTaxes = funcaoApoio.calculation_internal_matrix(mDistribution, mOffer, nColOtherTaxes)

# ============================================================================================
# Calculanting Coeficients without exports and Stock Variation
#  ============================================================================================
mDemandWithoutExport = np.copy(mDemandWithoutEstock)
# zerando coluna de exportações
mDemandWithoutExport[:, nColExport] = 0
# Diminuindo as exportações da Demanda Final e da Demanda Total
for p in range(nProduct):
    mDemandWithoutExport[p, nColTotalDemand] = mDemandWithoutEstock[p, nColTotalDemand] - mDemandWithoutEstock[p, nColExport]
    mDemandWithoutExport[p, nColFinalDemand] = mDemandWithoutEstock[p, nColFinalDemand] - mDemandWithoutEstock[p, nColExport]

mDistributionWithoutExport, mTotalConsumWithoutExport = funcaoApoio.calculation_distribution_matrix(mIntermConsum, mDemandWithoutExport)

# ============================================================================================
# Calculanting Arrays internally distributed by alphas without exports
#  ============================================================================================
mImport = funcaoApoio.calculation_internal_matrix(mDistributionWithoutExport,vImport,  nColImport)
mImportTax = funcaoApoio.calculation_internal_matrix(mDistributionWithoutExport, mOffer, nColImportTax)

# ============================================================================================
# Calculanting the Matrix of Consum with base Price
#  ============================================================================================
mTotalConsum =np.concatenate((mIntermConsum, mDemand), axis=1)
mConsumBasePrice = mTotalConsum - mMarginTrade - mMarginTransport - mIPI - mICMS - mOtherTaxes - mImport - mImportTax
nRow, nCol = np.shape(mConsumBasePrice)
# calculating totals of product (rows)
for r in range(nProduct):
    mConsumBasePrice[r, nSector] = sum(mConsumBasePrice[r, 0:nSector])
# calculating Total of totals of Cols ( CI and Final Demand
for c in range(nCol):
    mConsumBasePrice[nRow-1, c] = sum(mConsumBasePrice[0: nRow-1, c])

# Creating E Matrix with basic price
mE = mConsumBasePrice[:, nSector+1:nSector + nColsDemand + 2]
# ============================================================================================
# Calculanting Gross Value of Production by Product and check with production matrix
# ============================================================================================
mComplemen = np.zeros([nRowsAV+7, nCol], dtype=float)
mComplemen[0,:] = mImport[nProduct,:]
mComplemen[1,:] = mImportTax[nProduct,:]
mComplemen[2,:] = mIPI[nProduct,:]
mComplemen[3,:] = mICMS[nProduct,:]
mComplemen[4,:] = mOtherTaxes[nProduct,:]

mComplemen[6:nRowsAV+6,0:nSector+1] = mAddedValue
vNameComplemenBP = ['Importação', 'II', 'IPI', 'ICMS', 'OILL', 'CI'] + vNameAddedValue + ['Dif = CI + VA - Valor da produção']



# Calculanting Total of intermediate Consume with imports and all taxes
for c in range(nCol):
    mComplemen[5, c]=sum(mComplemen[0:5, c]) + mConsumBasePrice[nRow-1, c]

# Agregating VA into  Base price I-O Matrix
mBasePriceUses =np.concatenate((mConsumBasePrice, mComplemen ), axis=0)
nRowUses, nColUses = np.shape(mBasePriceUses)

# Checking consistence of I-O Matrix
for c in range(nColUses):
    mBasePriceUses[nRowUses-1, c] = mBasePriceUses[nRow+5, c] +   mBasePriceUses[nRow+6, c] -  mBasePriceUses[nRow+nRowsAV+6-2 , c]

# Creating D Matrix
mProductionTrans=mProduction.T
mD = np.zeros([ nSector+1, nProduct+1,], dtype=float)
for r  in range(nSector):
    for c in range(nProduct):
        if mProductionTrans[nSector, c]==0:
            mD[r, c] = mProductionTrans[r, c]
        else:
            mD [r, c] = mProductionTrans [r, c] / mProductionTrans [nSector, c]

# creating Matrix of National Coefficients - Bn Matrix
mBn = np.zeros([nProduct+1, nSector+1], dtype=float)
for r  in range(nProduct):
    for c in range(nSector):
        if ( mAddedValue[nRowTotalProduction, c] == 0):
            mBn[r, c] = 0.0
        else:
            mBn[r, c] = mConsumBasePrice [r, c] / mAddedValue [nRowTotalProduction, c]
#            mBn[r, c] = mConsumBasePrice[r, c] / mProductionTrans [nSector, c]

# Creating X Vector
vX = mAddedValue [nRowTotalProduction, 0:nSector]

# creating Matrix of imports Coefficients Bm Matrix
mBm = np.zeros([nProduct+1, nSector+1], dtype=float)
for r  in range(nProduct):
    for c in range(nSector):
        if ( vX[c] == 0):
            mBm[r, c] = 0.0
        else:
            mBm[r, c] = mImport [r, c] / vX[c]

# Creating A Matrix
mA = (np.dot(mD, mBn))

# Creating Z Matrix
mZ =np.zeros([ nSector+1, nSector+1], dtype=float)
for r  in range(nSector):
    for c in range(nSector):
        mZ[r, c] = mA[r, c] * vX[c]
        mZ[r, nSector] += mZ[r, c]
        mZ[nSector, c] += mZ[r, c]
mZ[nSector, nSector] = sum( mZ[:, nSector])
# Creating Y Matrix
mY = (np.dot(mD, mE))
for c in range(nColsDemand):
    mY[nSector, c] = sum(mY[:, c])

# calculanting Leontief Matrix  ( Sector x Sector )
mI = np.eye(nSector+1)
#mI = np.ones([ nSector+1, nSector+1,], dtype=float)
#mLeontief = ( mI - mA)**(-1)
mLeontief = np.linalg.inv(mI - mA)

mMIP =np.concatenate((mZ, mY), axis=1)
nRow, nCol = np.shape(mMIP)
mComplemen = np.zeros([nRowsAV+7, nCol], dtype=float)
mComplemen[0,:] = mImport[nProduct,:]
mComplemen[1,:] = mImportTax[nProduct,:]
mComplemen[2,:] = mIPI[nProduct,:]
mComplemen[3,:] = mICMS[nProduct,:]
mComplemen[4,:] = mOtherTaxes[nProduct,:]
mComplemen[6:nRowsAV+6,0:nSector+1] = mAddedValue
vNameComplemenMIP = ['Importação', 'II', 'IPI', 'ICMS', 'OILL', 'CI'] + vNameAddedValue\
                 + ['Dif = CI + VA - Valor da produção'] + ['Dif = Valor da produção - Demanda total']
# Calculanting Total of intermediate Consume with imports and all taxes
for c in range(nCol):
    mComplemen[5, c]=sum(mComplemen[0:5, c]) + mMIP[nRow-1, c]
# Agregating VA into  Base price I-O Matrix
mMIPGeral =np.concatenate((mMIP, mComplemen), axis=0)
nRowMIP, nColMIP = np.shape(mMIPGeral)


# Checking consistence of MIP
vDiff = np.zeros([nColMIP], dtype=float)
for c in range(nColMIP):
    mMIPGeral[nRowMIP-1, c] = mMIPGeral[nRow+5, c] +   mMIPGeral[nRow+6, c] - mMIPGeral[nRow+nRowsAV+6-2 , c]

for r in range(nSector+1):
    vDiff[r] = mMIPGeral[nRowMIP-3 , r] - mMIPGeral[r,nColMIP -1]
mMIPGeral = np.vstack((mMIPGeral, vDiff))

vGDP, vNameGDP, vNameColGDP =funcaoApoio.GDP_Calculation (mMIPGeral, nSector)


# ============================================================================================
# Writing Excel file with Sheets
# ============================================================================================
vDataSheet.append(mAddedValue)
vSheetName.append('VA')
vRowsLabel.append(vNameAddedValue)
vColsLabel.append(vNameSector)

vDataSheet.append(mDemand)
vSheetName.append('Demanda')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameDemand)

vDataSheet.append(mIntermConsum)
vSheetName.append('CI')
vRowsLabel.append(vNameProduct)
vNameCIDemand = vNameSector + vNameDemand
vColsLabel.append(vNameSector)

vDataSheet.append(mOffer)
vSheetName.append('Oferta')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameOffer)

vDataSheet.append(mProduction)
vSheetName.append('Producao')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameSector)

vDataSheet.append(vImport)
vSheetName.append('Importacao')
vRowsLabel.append(vNameProduct)
vNameImport = ['Importação']
vColsLabel.append(vNameImport)

vDataSheet.append(mDistribution)
vSheetName.append('Distribuição')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mMarginTrade)
vSheetName.append('MGC')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mMarginTransport)
vSheetName.append('MGT')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mIPI)
vSheetName.append('IPI')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mICMS)
vSheetName.append('ICMS')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mOtherTaxes)
vSheetName.append('OILL')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mDistributionWithoutExport)
vSheetName.append('Distribuição_2')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mImport)
vSheetName.append('Importação_2')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mImportTax)
vSheetName.append('II')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mBasePriceUses)
vSheetName.append('Usos PB')
vRowsLabel.append(vNameProduct + vNameComplemenBP)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(mBn)
vSheetName.append('Matriz_Bn')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameSector)

vDataSheet.append(mBm)
vSheetName.append('Matriz_Bm')
vRowsLabel.append(vNameProduct)
vColsLabel.append(vNameSector)

vDataSheet.append(mD)
vSheetName.append('Matriz_D')
vRowsLabel.append(vNameSector)
vColsLabel.append(vNameProduct)

vDataSheet.append(mA)
vSheetName.append('Matriz_A')
vRowsLabel.append(vNameSector)
vColsLabel.append(vNameSector)


vDataSheet.append(mZ)
vSheetName.append('Matriz_Z')
vRowsLabel.append(vNameSector)
vColsLabel.append(vNameSector)

vDataSheet.append(mY)
vSheetName.append('Matriz_Y')
vRowsLabel.append(vNameSector)
vColsLabel.append(vNameDemand)

vDataSheet.append(mLeontief)
vSheetName.append('Matriz_Leontief')
vRowsLabel.append(vNameSector)
vColsLabel.append(vNameSector)

vDataSheet.append(mMIPGeral)
vSheetName.append('MIP')
vRowsLabel.append(vNameSector + vNameComplemenMIP)
vColsLabel.append(vNameCIDemand)

vDataSheet.append(vGDP)
vSheetName.append('PIB')
vRowsLabel.append(vNameGDP)
vColsLabel.append(vNameColGDP)

funcaoApoio.write_data_excel(sFileSheet, vSheetName, vDataSheet, vRowsLabel, vColsLabel)

print("Terminou ")
sys.exit(0)

