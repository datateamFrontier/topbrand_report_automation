import pandas as pd
import numpy as np
import datetime
import seaborn as sb



def kota_multiindex():
    """
    Daftar regional dan kota untuk kolom berlevel
    """

    idx_kota = pd.MultiIndex.from_tuples([
        ('Jawa & Denpasar','Jabodetabek'),
        ('Jawa & Denpasar','Bandung'),
        ('Jawa & Denpasar','Semarang'),
        ('Jawa & Denpasar','Yogyakarta'),
        ('Jawa & Denpasar','Surabaya'),
        ('Jawa & Denpasar','Malang'),
        ('Jawa & Denpasar','Denpasar'),
        ('Sumatera','Medan'),
        ('Sumatera','Palembang'),
        ('Sumatera','Pekanbaru'),
        ('Kalimantan','Banjarmasin'),
        ('Kalimantan','Balikpapan'),
        ('Kalimantan','Samarinda'),
        ('Sulawesi','Makassar'),
        ('Sulawesi','Manado')
    ],names=['Kelompok Kota','Kota'])
    return idx_kota

def kota_urut():
    """
    Urutan kota dari Jawa, Sumatera, Kalimantan, lalu Sulawesi
    """

    kota = ['Jabodetabek','Bandung','Semarang','Yogyakarta','Surabaya','Malang','Denpasar',
            'Medan','Palembang','Pekanbaru',
            'Banjarmasin','Balikpapan','Samarinda',
            'Makassar','Manado']
    return kota


def ses_urut():
    """
    Urutan SES dari kecil ke besar
    """

    ses = ['SES D','SES C','SES B','SES A','SES A+']
    return ses

def pendidikan_urut():
    """
    Urutan pendidikan dari kecil ke besar
    """

    pendidikan = ['SD/sederajat', 'SMP/sederajat', 'SMA/sederajat', 'Akademi/ Diploma', 'Sarjana (S1/ S2/ S3)']
    return pendidikan

def conditional_formating(data, kriteria=None):
    cm = sb.light_palette("seagreen", as_cmap=True)
    if kriteria=='UN':
        id_min = -2
    elif (kriteria=='FI') or (kriteria=='LU') or (kriteria=='TOM') or (kriteria=='pekerjaan'):
        id_min = -1

    else:
        id_min = len(data)
        
    if kriteria == 'pekerjaan':
        colored_data = data.style.background_gradient(cmap=cm,
                                          subset=pd.IndexSlice[data.index[0]:data.index[9], :],
                                          vmin = min(data.iloc[:id_min,:].min()), vmax = max(data.iloc[:id_min,:].max()))
    elif kriteria=='competitor':
        colored_data = data.style.background_gradient(cmap=cm,
                                              subset=pd.IndexSlice[data.index[0]:data.index[10], :],
                                              vmin = min(data.iloc[:id_min,:].min()), vmax = max(data.iloc[:id_min,:].max())).highlight_null('royalblue')                                          
    else:                                      
        colored_data = data.style.background_gradient(cmap=cm,
                                          subset=pd.IndexSlice[data.index[0]:data.index[10], :],
                                          vmin = min(data.iloc[:id_min,:].min()), vmax = max(data.iloc[:id_min,:].max()))

    return colored_data


def sample_size(data, list_kolom):
    """
    Menghitung data sample per kota

    Parameter
    ---------
    data : data yang akan diolah
    list_kolom : array nama-nama kolom yang akan dipakai di tabel output

    Return
    ------
    sample_size : tabel jumlah n Sample per kota
    """

    sample_size = data[['panel','kota']].copy()
    sample_size = sample_size.pivot_table(index='kota', columns='panel', values ='panel',
                           aggfunc = len)

    sample_size['Total'] = sample_size.sum(axis=1)

    total_s_size = sample_size.sum(axis=0).to_frame().transpose()
    total_s_size.index = ['Total']

    sample_size = sample_size.reindex(index = kota_urut())

    sample_size = pd.concat([sample_size, total_s_size], axis=0)

    sample_size = sample_size.reset_index()
    sample_size.columns = ['Kota','Booster','Random','Total']
    sample_size = sample_size.reindex(columns=list_kolom)

    return sample_size

def hitung_nSampel(data_masukan, kolom1, kolom2 = None):
    """
    Menghitung nilai n Sampel berdasarkan suatu kategori (TOM, LU, FI)
    breakdown by karakteristik responden

    Parameter
    ---------
    data_masukan = data yang akan diolah
    kolom1 = kolom kriteria (TOM, LU, FI)
    kolom2 = kolom untuk breakdown (Kota, Usiar, Expandr, Pendidikan atau
    jika tidak di breakdown)
    """
    if kolom2 == None:
        ct = pd.crosstab(data_masukan[kolom1], data_masukan[kolom1])
        nSampel = np.diag(ct)
        df_nSampel = pd.DataFrame({kolom1: ct.index.to_list(),'nSampel': nSampel})
        df_nSampel = df_nSampel.set_index(kolom1)
        total = df_nSampel.sum(axis = 0).to_frame().transpose()
        total.index = ['Total']
        df_nSampel = pd.concat([df_nSampel, total], axis = 0)
        df_nSampel.index.rename(kolom1, inplace = True)

        return df_nSampel

    else:
        ct = pd.crosstab(data_masukan[kolom1],data_masukan[kolom2])
        ct.index = ct.index.astype(int)
        ct.index.rename(kolom1,inplace=True)

        total = ct.sum(axis=0).to_frame().transpose()
        total.index = ['Total']
        total.index.rename(kolom1,inplace=True)

        df_nSampel = pd.concat([ct, total],axis=0)
        df_nSampel['Total'] = df_nSampel.sum(axis=1)

        return(df_nSampel)


def buat_data_tracking(data, subkategori):
    """
    Membuat data tabel tracking brand selama 10 tahun terakhir

    Parameter
    ---------
    data : data yang akan diolah
    subkategori : subkategori yang akan dibuat laporan

    Return
    ------
    data : data brand beserta Top Brand Indeks-nya selama 10 tahun
    merek_tahun_ini : daftar brand yang akan di-tracking
    this_year : tahun saat laporan / data ini dibuat
    """

    today_ = datetime.datetime.now()
    this_year = today_.year

    data = data[(data.Subkategori.isin(subkategori)) & (data.Tahun>(this_year-10))]

    merek_tahun_ini = data[data.Tahun==this_year]
    merek_tahun_ini = merek_tahun_ini.sort_values(by='Top Brand Index',ascending=False)
    merek_tahun_ini = merek_tahun_ini.reset_index(drop=True)
    merek_tahun_ini = merek_tahun_ini.loc[0:5,'Merek'].to_list()

    data = data[data.Merek.isin(merek_tahun_ini)]

    return data, merek_tahun_ini, this_year


def top6_semuaElemen_DuaTahun(data, subkategori, tahun_ini):
    """
    Membuat data tabel nilai indeks TBI, TOM, FI, dan LU pada tahun ini dan sebelumnya
    untuk brand yang diamati

    Parameter
    ---------
    data : data yang akan diolah
    subkategori : subkategori yang akan dipakai
    tahun_ini : tahun saat data ini diolah / laporan dibuat

    Return
    ------
    data : data tabel masukan yang tinggal terdiri dari enam tahun saja
    data_sekarang : data tabel pengamatan untuk tahun sekarang
    data_kemarin : data tabel pengamatan untuk tahun kemarin
    """

    data = data[(data['Sub Kategori'].isin(subkategori)) & (data.Tahun>(tahun_ini-6))]

    data_sekarang = data[data.Tahun==tahun_ini]
    data_sekarang = data_sekarang.sort_values(by='TBI',ascending=False)
    data_sekarang = data_sekarang.reset_index(drop=True)
    data_sekarang = data_sekarang.loc[0:5, :]
    data_sekarang = data_sekarang.sort_values(by='Ranking',ascending=True)
    data_sekarang = data_sekarang.reset_index(drop=True)

    merek_top6 = list(data_sekarang.Brand[0:6].values)

    data_kemarin = data[data.Brand.isin(merek_top6)]
    data_kemarin = data_kemarin[data_kemarin.Tahun==(tahun_ini-1)]

    sorterIndex = dict(zip(merek_top6, range(len(merek_top6))))
    data_kemarin['Brand 1'] = data_kemarin['Brand'].map(sorterIndex)
    data_kemarin = data_kemarin.sort_values(by=['Brand 1'],ascending=True)
    data_kemarin = data_kemarin.drop(labels='Brand 1',axis=1)
    data_kemarin = data_kemarin.reset_index(drop=True)

    return data, data_sekarang, data_kemarin, merek_top6


def data_laporan(data, mascod_bedding,  top10_brand):
    """
    Memformat ulang data masukan dengan mengganti brand yang semula kode manjadi nama brand-nya

    Parameter
    ---------
    data : data yang akan diolah
    mascod_bedding : file excel daftar brand dengan kodenya
    top10_brand : daftar sepuluh besar brand

    Return
    ------
    df_laporan : data masukan yang sudah memiliki nama brand
    """

    df_laporan = data.copy()
    col_changed = ['TOM', 'UN1', 'UN2', 'UN3', 'UN4', 'UN5', 'LU', 'FI']
    n_brand = mascod_bedding.Brand.nunique()
    for i in range(n_brand):
        for col_j in col_changed:
            df_laporan.loc[df_laporan.loc[:,col_j] == mascod_bedding.loc[i,'Coding'], col_j] = mascod_bedding.loc[i,'Brand']
    for col_j in col_changed:
        df_laporan.loc[~(pd.isna(df_laporan[col_j])) & ~(df_laporan.loc[:,col_j].isin(top10_brand)), col_j] = 'Lainnya'

    return df_laporan


def data_ir(data):
    """
    Membuat tabel data perbandingan n Sampel dengan IR 'Ya' dan 'IR' tidak

    Parameter
    ---------
    data : data yang akan diolah

    Return
    ------
    tabel_totalBawah : tabel yang berisi jumlah n sampel dalam format angka biasa
    persen_tabel : tabel yang berisi jumlah n sampel dalam format persentase
    """

    crosstab = pd.crosstab(data['IR'],
                              data['kota'])
    crosstab = crosstab.reindex(columns = kota_urut())

    total_bawah = crosstab.sum(axis=0).to_frame().transpose()
    total_bawah.index = ['Total']
    total_bawah.index.rename('IR',inplace=True)

    tabel_totalBawah = pd.concat([crosstab,
                            total_bawah],
                           axis=0)
    tabel_totalBawah['Total'] = tabel_totalBawah.sum(axis=1)
    tabel_totalBawah = tabel_totalBawah.transpose()
    tabel_totalBawah.reset_index(inplace=True)
    tabel_totalBawah.columns.name=None
    tabel_totalBawah.set_index('kota',inplace=True)
    tabel_totalBawah.reset_index(inplace=True)

    persen_tabel = tabel_totalBawah.copy()
    persen_tabel.Tidak = (persen_tabel.Tidak/persen_tabel.Total)*100
    persen_tabel.Tidak = round(persen_tabel.Tidak, 1)
    persen_tabel.Ya = (persen_tabel.Ya/persen_tabel.Total)*100
    persen_tabel.Ya = round(persen_tabel.Ya, 1)
    persen_tabel.reset_index(drop=True, inplace=True)

    return tabel_totalBawah, persen_tabel


def profil_responden(data, kolom1, kolom2, sorter1, sorter2 = None):
    """
    Membuat data tabel profil responden

    Parameter
    ---------
    data : data yang akan diolah
    kolom1 : LU atau FI
    kolom2 : kolom untuk breakdown kolom1
    sorter1 : data series untuk mengurutkan brand
    sorter2 : list ubtuk sorter karakter responden
    Return
    ------
    total_saja : tabel yang berisi total n sampel masing-masing kategori kolom
    tabel_totalBawah : tabel yang berisi n sampel masing-masing kategori kolom untuk seluruh Brand
    """

    crosstab = pd.crosstab(data[kolom1],
                              data[kolom2])
    crosstab.index.rename(kolom1,inplace=True)

    if sorter2 != None:
        crosstab = crosstab.reindex(columns = sorter2)

    total_bawah = crosstab.sum(axis=0).to_frame().transpose()
    total_bawah.index = ['Total']
    total_bawah.index.rename(kolom1,inplace=True)

    tabel_totalBawah = pd.concat([crosstab,
                            total_bawah],
                           axis=0)
    tabel_totalBawah['Total'] = tabel_totalBawah.sum(axis=1)
    tabel_totalBawah['sorter'] = sorter1.to_list()
    tabel_totalBawah.sort_values(by='sorter',inplace=True)
    tabel_totalBawah.drop(columns='sorter',axis=1,inplace=True)
    tabel_totalBawah.reset_index(inplace=True)
    tabel_totalBawah.columns = ['Brand']+tabel_totalBawah.columns.to_list()[1:]

    total_saja = tabel_totalBawah.iloc[-1:,1:-1].copy()
    total_saja = round(total_saja/tabel_totalBawah.iloc[-1:,-1:].values*100,1)
    total_saja.reset_index(drop=True,inplace=True)

    return total_saja,tabel_totalBawah

def presentase_profil(data_masukan):
    """
    Menghitung presentase profil responden dari data count
    Parameter
    ---------
    data_masukan : data yang akan diolah
    ---------
    """
    
    dt_pekerjaan = data_masukan.copy()
    dt_pekerjaan.set_index('Brand',inplace=True)
    dt_pekerjaan = dt_pekerjaan.transpose()
    temp1 = dt_pekerjaan.iloc[:-1,:].sort_values(by='Total',ascending=False)
    temp2 = dt_pekerjaan.iloc[-1:,:]

    dt_pekerjaan = pd.concat([temp1[:10], temp2],axis=0)
    dt_pekerjaan = dt_pekerjaan.drop(labels = 'Total', axis = 1)
    dt_pekerjaan = round(100*(dt_pekerjaan/dt_pekerjaan.iloc[-1,:]),1)
    dt_pekerjaan.reset_index(inplace=True)
    dt_pekerjaan.columns = ['Pekerjaan']+dt_pekerjaan.columns[1:].to_list()
    dt_pekerjaan = conditional_formating(dt_pekerjaan.set_index('Pekerjaan'), 'pekerjaan')
    
    return(dt_pekerjaan)

def hitung_nilai(kriteria, data_used, indeks_brand, non_user = False):
    """
    Menghitung nilai indeks kriteria (TOM, LU, atau FI) untuk top Brand.

    Parameter
    ---------
    kriteria : TOM, LU, atau FI
    data_used : data yang akan diolah
    indeks_brand : data series untuk mengurutkan brand
    non_user : Jika TRUE, akan menghitung data yang memiliki nilai IR = 'Tidak'

    Return
    ------
    df_kriteria_sort : tabel nilai indeks brand sesuai kriteria masukan
    sum_bobot_ : tabel nilai total bobot brand
    """

    if (non_user == True):
        data_used = data_used[data_used.IR == 'Tidak']

    # Menghitung sum bobot cross tab by tom
    sum_bobot_ = data_used[['bobot',kriteria]].groupby(kriteria).sum()
    sum_bobot_.reset_index(inplace = True)
    total_bobot_ = pd.DataFrame({kriteria:['Total'], 'bobot':[sum(sum_bobot_.bobot)]})
    sum_bobot_ = pd.concat([sum_bobot_, total_bobot_], axis = 0)
    sum_bobot_.reset_index(drop = True, inplace = True)

    # Menghitung kriteria yang dicari
    df_kriteria = sum_bobot_.copy()
    df_kriteria.bobot = (df_kriteria.bobot/df_kriteria.bobot[len(df_kriteria)-1])*100
    if (kriteria == 'TOM'):
        nama_kolom = 'Mind Share'
    elif (kriteria == 'LU'):
        nama_kolom = 'Market Share'
    elif (kriteria == 'FI'):
        nama_kolom = 'Commitment Share'
    else:
        nama_kolom = df_kriteria.columns[1]

    df_kriteria.columns = ['Brand', nama_kolom]
    df_kriteria['sorter'] = indeks_brand
    df_kriteria = df_kriteria.sort_values(by='sorter')
    df_kriteria = df_kriteria.drop('sorter', axis = 1)
    df_kriteria = df_kriteria[df_kriteria.Brand != 'Total']
    df_kriteria.reset_index(drop = True, inplace = True)

    df_kriteria[nama_kolom] = round(df_kriteria[nama_kolom], 1)

    return df_kriteria, sum_bobot_


def hitung_gap(df_kriteria_sort, top10_brand):
    """
    Menghitung gap indeks tiap brand terhadap indeks brand tertinggi

    Parameter
    ---------
    df_kriteria_sort : data masukan yang akan dicari gap indeks tiap brand-nya

    Return
    ------
    gap_kriteria : tabel yang berisi nilai selisih indeks tiap brand dengan indeks tertinggi
    """

    nama_kolom = df_kriteria_sort.columns[1]

    # Menghitung gap from the highest
    max_kriteria = max(df_kriteria_sort.loc[df_kriteria_sort.Brand.isin(top10_brand),nama_kolom])
    gap_kriteria = df_kriteria_sort.copy()
    gap_kriteria.columns = ['Brand', 'Gap']
    gap_kriteria.Gap = abs(gap_kriteria.Gap - max_kriteria)
    gap_kriteria.reset_index(drop = True, inplace = True)

    # Pembulatan
    gap_kriteria.Gap = round(gap_kriteria.Gap, 1)
    return gap_kriteria


def hitung_nilai_crosstab(kriteria, by, data_used, indeks_brand, indeks_kolom = [], non_user = False,
                         dt_tom = None):
    """
    Membuat tabel crosstab untuk pemetaan top brand

    Parameter
    ---------
    kriteria : diisi TOM, UN, LU, atau FI
    by : kolom tabel crosstab terhadap tiap Brand
    data_used : data yang digunakan untuk pengolahan crosstab
    indeks_brand : daftar brand untuk mengurutkan brand
    indeks_kolom : kolom custom yang digunakan untuk output data
    non_user: jika TRUE, data laporan yang diambil adalah yang IR='Tidak'
    dt_tom : data laporan keseluruhan yang digunakan

    Return
    ------
    df_kriteria_by : tabel crosstab hasil perhitungan
    """

    if (non_user == True):
        data_used = data_used[data_used.IR == 'Tidak']

    # Menghitung sum bobot cross tab by kriteria dan karakter responden terpilih
    sum_bobot_ = pd.crosstab(index = data_used[kriteria], columns = data_used[by], values = data_used.bobot, aggfunc = sum)
    sum_bobot_.fillna(0, inplace = True)
    sum_bobot_.index.rename(kriteria, inplace=True)

    # Menghitung total bobot per kota
    total_bobot_ = sum_bobot_.sum(axis = 0).to_frame().transpose()
    total_bobot_.index = ['Total']
    total_bobot_.index.rename(kriteria, inplace=True)
    sum_bobot_ = pd.concat([sum_bobot_, total_bobot_], axis = 0)

    # Menghitung nilai kriteria per by
    df_kriteria_by = sum_bobot_.copy()

    if kriteria=='UN':
        sum_bobot_1 = pd.crosstab(index = dt_tom['TOM'], columns = dt_tom[by], values = dt_tom.bobot, aggfunc = sum)
        sum_bobot_1.fillna(0, inplace = True)
        sum_bobot_1.index.rename(kriteria, inplace=True)
        total_bobot_1 = sum_bobot_1.sum(axis = 0).to_frame().transpose()
        total_bobot_1.index = ['Total']
        total_bobot_1.index.rename(kriteria, inplace=True)
        sum_bobot_1 = pd.concat([sum_bobot_1, total_bobot_1], axis = 0)

        df_tom = sum_bobot_1.copy()
        col_by = df_tom.columns

        for col_by_i in col_by:
            df_kriteria_by.loc[:,col_by_i] = ((df_kriteria_by.loc[:,col_by_i]+df_tom.loc[:,col_by_i])/df_tom.loc['Total', col_by_i])*100
    else:
        col_by = df_kriteria_by.columns
        for col_by_i in col_by:
            df_kriteria_by.loc[:,col_by_i] = (df_kriteria_by.loc[:,col_by_i]/df_kriteria_by.loc['Total', col_by_i])*100

    df_kriteria_by.reset_index(inplace = True)
    df_kriteria_by['sorter'] = indeks_brand
    df_kriteria_by = df_kriteria_by.sort_values(by = 'sorter', axis = 0)
    df_kriteria_by = df_kriteria_by.drop('sorter', axis = 1)

    kol_ = df_kriteria_by.columns[0]
    df_kriteria_by = df_kriteria_by[df_kriteria_by[kol_] != 'Total']
    df_kriteria_by.columns.name = None

    if kriteria == 'UN':
        df_kriteria_by.set_index(kriteria, inplace = True)
        tot_multirespon = df_kriteria_by.sum(axis = 0).to_frame().transpose()
        tot_multirespon.index = ['Total multirespon']
        df_kriteria_by = pd.concat([df_kriteria_by, tot_multirespon], axis = 0)
        df_kriteria_by.index.rename(kriteria, inplace = True)
        df_kriteria_by.reset_index(inplace = True)

    # Sorting kolom
    if by == 'kota' or by=='expandr':
        df_kriteria_by.set_index(kriteria, inplace = True)
        df_kriteria_by = df_kriteria_by.reindex(columns= indeks_kolom)
        df_kriteria_by.reset_index(inplace = True)
    else:
        df_kriteria_by.reset_index(drop=True, inplace=True)

    df_kriteria_by[col_by] = round(df_kriteria_by[col_by], 1)

    return df_kriteria_by


def fungsi_by(kriteria, kolom, tabel_indeks, data_masukan):
    """
    Mengubah kolom brand menjadi indeks dan menambah baris data n sample

    Parameter
    ---------
    kriteria : diisi TOM, LU, atau FI
    kolom : nama kolom untuk crosstabulasi
    tabel_indeks : tabel indeks yang akan diolah
    data_masukan : tabel data acuan keseluruhan

    Return
    ------
    tabel_indeks : crosstabulasi dengan brand sebagai indeksnya
    """

    tabel_indeks.columns = ['Brand']+tabel_indeks.columns[1:].to_list()

    tb_crosstab = pd.crosstab(data_masukan[kriteria],
                              data_masukan[kolom])
    tb_crosstab.index.rename(kriteria,inplace=True)

    total_tb_kriteria = tb_crosstab.sum(axis=0).to_frame().transpose()
    total_tb_kriteria.index = ['Total']
    total_tb_kriteria.index.rename(kriteria,inplace=True)

    tb_kriteria = pd.concat([tb_crosstab,
                            total_tb_kriteria],
                           axis=0)
    tb_kriteria['Total'] = tb_kriteria.sum(axis=1)

    tabel_indeks = tabel_indeks.append(pd.Series(dtype='object'), ignore_index=True)
    n_sample = tb_kriteria.iloc[-1:,:].reset_index(drop=True)

    for kt in tabel_indeks.columns:
        if kt=='Brand':
            tabel_indeks.loc[len(tabel_indeks)-1:,
                                'Brand']='n Sample'
        else:
            tabel_indeks.loc[len(tabel_indeks)-1:,
                                kt]=n_sample[kt].values[0]

    tabel_indeks.set_index('Brand', inplace=True)
    if kolom=='kota':
        tabel_indeks.columns = kota_multiindex()
    tabel_indeks = conditional_formating(tabel_indeks, kriteria)

    return tabel_indeks


def hitung_tbi(tom, lu, fi):
    """
    Menghitung nilai TBI tiap brand menggunakan tabel TOM, LU, dan FI

    Parameter
    ---------
    tom : tabel data indeks TOM
    lu : tabel data indeks LU
    fi : tabel data indeks FI

    Return
    ------
    tbi : tabel dengan nilai indeks TBI per brand
    """

    dt_tbi = pd.DataFrame(columns=['Brand','TBI'],index=range(len(tom)))
    for jml_brand in range(len(dt_tbi)):
        dt_tbi.loc[jml_brand, 'Brand'] = tom.loc[jml_brand, 'Brand']
        dt_tbi.loc[jml_brand, 'TBI'] = float(0.4*tom.loc[jml_brand,'Mind Share'] +
                                    0.3*lu.loc[jml_brand,'Market Share'] +
                                    0.3*fi.loc[jml_brand,'Commitment Share'])
    dt_tbi['TBI'] = dt_tbi['TBI'].astype('float64')
    dt_tbi['TBI'] = dt_tbi['TBI'].round(decimals=1)

    return dt_tbi


def crosstab_tbi(tom, lu, fi):
    """
    Membuat crosstabulasi tabel indeks tbi

    Parameter
    ---------
    tom : tabel data indeks TOM
    lu : tabel data indeks LU
    fi : tabel data indeks FI

    Return
    ------
    tbi : tabel dengan crosstabulasi indeks TBI per brand
    """

    dt_tbi_by = pd.DataFrame(columns=tom.columns,index=tom.index)
    for brand in dt_tbi_by.index:
        for kolom in dt_tbi_by.columns:
            dt_tbi_by.loc[brand, kolom] = float(0.4*tom.loc[brand,kolom] +
                                         0.3*lu.loc[brand,kolom]+
                                         0.3*fi.loc[brand,kolom])
            dt_tbi_by.loc[brand, kolom] = round(dt_tbi_by.loc[brand,kolom],1)
    dt_tbi_by = dt_tbi_by.drop(labels = 'n Sample', axis = 0)

    return dt_tbi_by


def get_unaided(data):
    """
    Mendapatkan daftar kolom Unaided

    Parameter
    ---------
    data : data masukan untuk pembuatan laporan

    Return
    ------
    list_un : daftar kolom unaided
    """

    list_un = []
    j=1
    for i in range(len(data.columns)):
        cln = data.columns[i]
        if cln == 'UN%d'%(j):
            list_un.append(cln)
            j+=1
    return list_un


def get_dataframe_unaided(data):
    """
    Membuat data tabel khusus untuk data Unaided

    Parameter
    ---------
    data : data masukan untuk pembuatan laporan

    Return
    ------
    dt : tabel data khusus Unaided
    """

    new_kolom = ['no_entry', 'panel', 'kota', 'bobot', 'sex', 'usia', 'usiar', 'UN',
                'didik', 'kerja', 'expand', 'expandr']
    dt = pd.DataFrame(columns=new_kolom)
    list_kolom_un = get_unaided(data)
    for ke in range(len(list_kolom_un)):
        kolom = ['no_entry', 'panel', 'kota', 'bobot', 'sex', 'usia', 'usiar', list_kolom_un[ke],
                 'didik', 'kerja', 'expand', 'expandr']
        temp_dt = data[kolom].copy()
        temp_dt.columns = new_kolom

        dt = pd.concat([dt, temp_dt],axis=0, ignore_index=True)
    dt.reset_index(drop=True)
    return dt


def hitung_unaided(data_unaided, tabel_bobot_tom, sorter_brand):
    """
    Menghitung indeks untuk brand di Unaided

    Parameter
    ---------
    data_unaided : data masukan laporan khusus bagian unaided
    tabel_bobot_tom : tabel bobot hasil dari perhitungan brand TOM
    sorter_brand : daftar nomor untuk mengurutkan brand

    Return
    ------

    """

    dt_unaided = data_unaided[['UN','bobot']].groupby(by='UN').sum().reset_index()
    dt_unaided.columns = ['Brand','Unaided']

    dt_unaided['Unaided'] = (dt_unaided['Unaided']+tabel_bobot_tom['bobot'][:-1])/tabel_bobot_tom['bobot'][len(tabel_bobot_tom)-1]*100
    dt_unaided['Unaided'] = dt_unaided['Unaided'].astype('float64').round(1)

    dt_unaided['sorter'] = sorter_brand
    dt_unaided.sort_values(by='sorter',inplace=True)
    dt_unaided.drop(columns='sorter',inplace=True)
    dt_unaided.reset_index(drop=True, inplace=True)

    return dt_unaided


def hitung_brand_switch(data_used, list_sorted_brand, sorter_brand):
    """
    Membuat tabel brand switching analysis

    Parameter
    ---------
    data_used : data tabel yang digunakan untuk mengolah laporan
    list_sorted_brand : data brand yang digunakan untuk sorting (brand dan urutan ranking-nya)
    sorter_brand : data brand yang digunakan untuk sorting (nomornya saja)

    Return
    ------
    tab_switching : tabel nilai brand switching
    """

    # Menghitung sum bobot cross tab by LU dan FI
    bobot_switching = pd.crosstab(index = data_used['LU'], columns = data_used['FI'], values = data_used.bobot, aggfunc = sum)
    bobot_switching.fillna(0, inplace = True)

    # Menghitung total bobot per LU
    tot_bobot_switching = bobot_switching.sum(axis = 1).to_frame()
    tot_bobot_switching.columns = ['Total']
    bobot_switching_tot = pd.concat([bobot_switching, tot_bobot_switching], axis = 1)

    ##  Menghitung bobot loyalist, switching in, switching out
    tab_bobot_switching = pd.DataFrame(index = bobot_switching.index, columns = ['Loyalist', 'Switching out', 'Switching in', 'Total'])

    # Loyalist
    tab_bobot_switching['Loyalist'] = np.diag(bobot_switching)

    # Switcing in dan switching out
    nrow = len(bobot_switching.index)
    bobot_switching_0 = bobot_switching.copy()
    for i in range(nrow):
        bobot_switching_0.iloc[i,i] = 0

    tab_bobot_switching['Switching out'] = bobot_switching_0.sum(axis = 1).to_frame()
    tab_bobot_switching['Switching in'] = bobot_switching_0.sum(axis = 0).to_frame()
    tab_bobot_switching['Total'] = bobot_switching_tot['Total']

    # Mengurutkan berdasarkan tbi
    brand_sorted_ = list_sorted_brand.sort_values(by='Sorting').reset_index(drop=True)['Brand'][:len(list_sorted_brand)-1].to_list()
    tab_bobot_switching_sort = tab_bobot_switching.reindex(index = brand_sorted_)

    # Menghitung nilai loyalist, switching in, switching out
    tab_switching = pd.DataFrame(index = tab_bobot_switching_sort.index, columns = ['Loyalist', 'Switching out', 'Switching in', 'Net switching', 'Actual LU', 'Prediksi LU'])
    tab_switching['Loyalist'] = (tab_bobot_switching_sort['Loyalist']/tab_bobot_switching_sort['Total'])*100
    tab_switching['Switching out'] = (tab_bobot_switching_sort['Switching out']/tab_bobot_switching_sort['Total'])*100
    tab_switching['Switching in'] = (tab_bobot_switching_sort['Switching in']/tab_bobot_switching_sort['Total'])*100
    tab_switching['Net switching'] = tab_switching['Switching in'] - tab_switching['Switching out']

    actual_lu,_ = hitung_nilai('LU', data_used, sorter_brand)
    actual_lu.set_index('Brand', inplace = True)

    tab_switching['Actual LU'] = actual_lu
    tab_switching['Prediksi LU'] = tab_switching['Actual LU'] + tab_switching['Actual LU']*tab_switching['Net switching']/100

    # Pembulatan
    tab_switching = round(tab_switching, 1)
    tab_switching.drop(labels = 'Lainnya', axis = 0, inplace = True)
    tab_switching.reset_index(inplace = True)

    tab_switching.columns = ['Brand']+tab_switching.columns[1:].to_list()

    return tab_switching


def n_sample_user(data_user, sorter_brand):
    """
    Membuat tabel data n Sample berdasar IR=Ya

    Parameter
    ---------
    data_user : data tabel olah laporan dengan IR=Ya
    sorter_brand : data brand untuk mengurutkan index

    Return
    ------
    n_sample : tabel nilai n Sample
    """
    n_sample = data_user['LU'].value_counts().to_frame()
    n_sample.columns = ['n Sample']
    n_sample = n_sample.reindex(index=sorter_brand.sort_values(by='Sorting').reset_index(drop=True)['Brand'][:len(sorter_brand)-1].to_list())
    n_sample.reset_index(inplace=True)
    n_sample.columns = ['Brand','n Sample']

    return n_sample


def tabel_brand_switch(data_used, list_sorted_brand, n_sample):
    """
    Membuat tabel brand switch dalam bentuk crosstab FI--LU

    Parameter
    ---------
    data_used : data laporan yang akan diolah
    list_sorted_brand : data untuk mengurutkan daftar brand
    n_sample : tabel data berisi n Sample

    Return
    ------
    tab_switching_sort_t : tabel luaran berbentuk crosstab
    """
    # Menghitung sum bobot cross tab by LU dan FI
    bobot_switching = pd.crosstab(index = data_used['LU'], columns = data_used['FI'], values = data_used.bobot, aggfunc = sum)
    bobot_switching.fillna(0, inplace = True)

    # Menghitung total bobot per LU
    tot_bobot_switching = bobot_switching.sum(axis = 1).to_frame()
    tot_bobot_switching.columns = ['Total']
    bobot_switching_tot = pd.concat([bobot_switching, tot_bobot_switching], axis = 1)

    # ## Menghitung Presentase
    tab_switching = bobot_switching.copy()
    brand = bobot_switching.columns
    for brand_i in brand:
        tab_switching.loc[brand_i ,:] = (tab_switching.loc[brand_i ,:]/bobot_switching_tot.loc[brand_i, 'Total'])*100

    ## Mengurutkan berdasarkan tbi
    brand_sorted_ = list_sorted_brand.Brand.to_list()
    tab_switching_sort = tab_switching.reindex(index = brand_sorted_, columns = brand_sorted_)

    ## Transpose
    tab_switching_sort_t = tab_switching_sort.transpose()
    tab_switching_sort_t.index.rename('FI', inplace = True)

    ## Pembulatan
    tab_switching_sort_t = round(tab_switching_sort_t, 1)

    tab_switching_sort_t = pd.concat([tab_switching_sort_t, n_sample.set_index('Brand').transpose()], axis=0)
    tab_switching_sort_t.index.name='FI'
    tab_switching_sort_t.columns.name=''
    tab_switching_sort_t = pd.concat([tab_switching_sort_t], keys=['LU'], axis=1)

    #tab_switching_sort_t.drop(labels = 'Lainnya', axis = 0, inplace = True)
    #tab_switching_sort_t = tab_switching_sort_t.iloc[:, :len(tab_switching_sort_t)-1]
    return tab_switching_sort_t


# Fungsi perhitungan tabel convertion rate
def hitung_conv_rate(data_used, list_sorted_brand):
    """
    Membuat tabel berisi nilai conversion rate

    Parameter
    ---------
    data_used : data laporan yang digunakan untuk mengolah
    list_sorted_brand : tabel yang digunakan untuk mengurutkan brand

    Return
    ------
    tab_conv_rate : tabel luaran data conversion rate
    """
    ## Menghitung bobot tom
    bobot_tom = data_used[['bobot','TOM']].groupby('TOM').sum()
    bobot_tom.fillna(0, inplace = True)
    bobot_tom.index.rename('Brand', inplace=True)
    # display(bobot_tom)

    ## Menghitung bobot CR
    data_cr = data_used.loc[data_used['TOM']==data_used['LU'], ['TOM', 'bobot']]
    bobot_cr = data_cr.groupby('TOM').sum()
    bobot_cr.index.rename('Brand', inplace=True)

    ## Menghitung bobot UOB
    data_ir_ya = data_used[data_used['IR'] == 'Ya']
    data_uob = data_ir_ya.loc[data_ir_ya['TOM']!=data_ir_ya['LU'] , ['LU','TOM', 'bobot']]
    bobot_uob = data_uob.groupby('TOM').sum()
    bobot_uob.index.rename('Brand', inplace=True)

    ## Menghitung bobot ABNU
    data_abnu = data_used.loc[data_used['IR'] == 'Tidak', ['TOM', 'bobot']]
    bobot_abnu = data_abnu.groupby('TOM').sum()
    bobot_abnu.index.rename('Brand', inplace=True)

    ## Menghitung bobot NABU
    data_nabu = data_ir_ya.loc[data_ir_ya['TOM']!=data_ir_ya['LU'] , ['LU', 'bobot']]
    bobot_nabu = data_nabu.groupby('LU').sum()
    bobot_nabu.index.rename('Brand', inplace=True)

    ## Menyusun tabel bobot perhitungan convertion rate
    bobot_conv_rate = pd.DataFrame(index = bobot_nabu.index, columns = ['TOM', 'CR', 'UOB', 'ABNU', 'NABU'])
    bobot_conv_rate['TOM'] = bobot_tom
    bobot_conv_rate['CR'] = bobot_cr
    bobot_conv_rate['UOB'] = bobot_uob
    bobot_conv_rate['ABNU'] = bobot_abnu
    bobot_conv_rate['NABU'] = bobot_nabu

    ## Mengurutkan brand berdasarkan tbi'
    brand_sorted_ = list_sorted_brand.Brand.to_list()
    bobot_conv_rate_sort = bobot_conv_rate.reindex(index = brand_sorted_)

    ## Menghitung tabel nilai convertion rate
    tot_bobot_tom = sum(bobot_conv_rate_sort.TOM)
    tab_conv_rate = bobot_conv_rate_sort.copy()
    tab_conv_rate['TOM'] = (bobot_conv_rate_sort['TOM']/tot_bobot_tom)*100
    tab_conv_rate['CR'] = (bobot_conv_rate_sort['CR']/bobot_conv_rate_sort['TOM'])*100
    tab_conv_rate['UOB'] = (bobot_conv_rate_sort['UOB']/bobot_conv_rate_sort['TOM'])*100
    tab_conv_rate['ABNU'] = (bobot_conv_rate_sort['ABNU']/bobot_conv_rate_sort['TOM'])*100
    tab_conv_rate['NABU'] = (bobot_conv_rate_sort['NABU']/bobot_conv_rate_sort['TOM'])*100

    ## Pembulatan
    tab_conv_rate = round(tab_conv_rate, 1)
    tab_conv_rate.drop(labels = 'Lainnya', axis = 0, inplace = True)
    tab_conv_rate = tab_conv_rate.reset_index()

    return tab_conv_rate


def brand_diagnostic(tabel_tom, tabel_lu, tabel_fi):
    """
    Membuat data tabel untuk brand diagnostic

    Parameter
    ---------
    tabel_tom : data indeks brand di TOM
    tabel_lu : data indeks brand di LU
    tabel_fi : data indeks brand di FI

    Return
    ------
    dt_diagnostic : tabel berisi nilai brand diagnostic
    """

    df_tlf = pd.concat([tabel_tom, tabel_lu, tabel_fi], axis=1)
    df_tlf = df_tlf.loc[:,~df_tlf.columns.duplicated()]
    df_tlf.columns = ['Brand','TOM','LU','FI']

    df_bdiagnostic = pd.DataFrame(columns=['Brand','TOM/LU','FI/LU','TOM/FI'])
    df_bdiagnostic['Brand'] = df_tlf['Brand']
    df_bdiagnostic['TOM/LU'] = round(df_tlf['TOM']/df_tlf['LU'],3)
    df_bdiagnostic['FI/LU'] = round(df_tlf['FI']/df_tlf['LU'],3)
    df_bdiagnostic['TOM/FI'] = round(df_tlf['TOM']/df_tlf['FI'],3)

    dt_diagnostic = pd.concat([df_tlf, df_bdiagnostic], axis=1)
    dt_diagnostic = dt_diagnostic.loc[:,~dt_diagnostic.columns.duplicated()]

    dt_diagnostic = dt_diagnostic.loc[dt_diagnostic.Brand != 'Lainnya', :]

    return dt_diagnostic


def hitung_bobot_cl(data_masukan, sorter_brand):
    """
    Menghitung bobot overall competitor landscape

    Parameter
    ---------
    data_masukan : data yang akan diolah / digunakan untuk mencari bobot
    sorter_brand : data untuk mengurutkan brand

    Return
    ------
    bobot_total_cl : data bobot overall per brand
    """

    cl_awal = pd.crosstab(index = data_masukan['LU'], columns = data_masukan['kota'], values = data_masukan.bobot, aggfunc = sum)
    cl_awal.fillna(0, inplace = True)
    cl_awal.index.rename('LU', inplace=True)
    cl_awal.reset_index(inplace=True)
    cl_awal.set_index(['LU'],inplace=True)
    cl_awal= cl_awal.reindex(sorter_brand['Brand'])
    cl_awal = cl_awal.reindex(columns=kota_urut())
    cl_awal['Total'] = cl_awal.sum(axis=1)

    cl_total = cl_awal.sum(axis = 0).to_frame().transpose()
    cl_total.index = ['Total']
    cl_total.index.rename('LU', inplace=True)

    cl_awal = pd.concat([cl_awal, cl_total], axis = 0)
    cl_awal.index.rename('Brand',inplace=True)

    cl_kesamping = cl_awal.iloc[:-1,:-1].copy()
    for idx in cl_kesamping.index:
        cl_kesamping.loc[idx,:] = cl_kesamping.loc[idx,:]/cl_awal.loc[idx,'Total']
    cl_kesamping.index.name = ''

    cl_kebawah = cl_awal.iloc[:-1,:-1].copy()
    for clm in cl_kebawah.columns:
        cl_kebawah.loc[:,clm] = cl_kebawah.loc[:,clm]/cl_awal.loc['Total',clm]
    cl_kebawah = cl_kebawah.transpose()
    cl_kebawah.columns.name = ''

    competitor_landscape = []

    kt = cl_kesamping.columns.to_list()
    brd = cl_kesamping.index.to_list()

    for kt_ in range(len(kt)):
        tb_cl = pd.DataFrame(index=cl_kesamping.index.copy(), columns=cl_kebawah.columns.copy())
        for brd1 in brd:
            for brd2 in brd:
                if brd1==brd2:
                    continue
                tb_cl.loc[brd1, brd2] = cl_kesamping.loc[brd1,kt[kt_]] * cl_kebawah.loc[kt[kt_], brd2]
        tb_cl.index.name = kt[kt_]
        competitor_landscape.append(tb_cl)


    competitor_all = pd.DataFrame(columns=cl_kebawah.columns,
                                 index=cl_kesamping.index)
    for brd1 in brd:
        for brd2 in brd:
            if brd1==brd2:
                continue
            nilai_akumulasi = 0
            for cl in range(len(competitor_landscape)):
                nilai_akumulasi += competitor_landscape[cl].loc[brd1, brd2]
            competitor_all.loc[brd1, brd2] = nilai_akumulasi
    bobot_total_cl = competitor_all.sum(axis=1)

    return bobot_total_cl, competitor_all, competitor_landscape


def hitung_competitor_landscape(nama_tabel, tabel_masukan,
                               bobot_pembagi):
    """
    Membuat tabel crosstab competitor landscape

    Parameter
    ---------
    nama_tabel : nama untuk tabel competitor landscape-nya
    tabel_masukan : tabel yang digunakan untuk menghitung
    bobot_pembagi : digunakan untuk menghitung nilai indeks

    Return
    ------
    tabel_output : data luaran dalam bentuk tabel
    """

    tabel_output = pd.DataFrame(index = tabel_masukan.index.copy(), columns=tabel_masukan.columns.copy())
    for index in tabel_output.index:
        for column in tabel_output.columns:
            if index==column:
                continue
            tabel_output.loc[index, column] = round(tabel_masukan.loc[index,column]/bobot_pembagi[index],3)
    tabel_output.index.name = ''
    tabel_output = pd.concat([tabel_output], keys=['Focal Firm'],names=[nama_tabel])
    tabel_output = pd.concat([tabel_output], keys=['Competitors'], axis=1)

    return tabel_output

# ----------------------------------------------------------------------------------------------------------------------------

def filter_tambahan(data_used, kode_awal, incl_kota = True, kolom_kota = 'Kota_1_0'):
    """
    Mem-filter kolom yang akan digunakan untuk data tambahan

    Parameter
    ---------
    data_used : data yang digunakan untuk membuat general information
    kode_awal : awalan nama kolom yang ingin di-filter
    incl_kota : untuk memgganti nama kolom menjadi kota
    kolom_kota : nama kolom kota

    Return
    ------
    data_tambahan_ : data hasil filter
    """

    kolom_data = data_used.columns.to_list()
    kolom_tambahan = []
    for i in range(len(kolom_data)):
        if kode_awal in kolom_data[i]:
            kolom_tamb_i = kolom_data[i]
            kolom_tambahan.append(kolom_tamb_i)

    if kode_awal[:2] == 'tv':
        kol_tv = []
        for kol in kolom_tambahan:
            if kol[:2] == 'tv':
                kol_tv.append(kol)
        kolom_tambahan = kol_tv

    data_tambahan_ = data_used[kolom_tambahan]
    if incl_kota == True:
        data_tambahan_ = data_used[[kolom_kota]+kolom_tambahan]
    return data_tambahan_


def hitung_tambahan(data_used, kode_awal, mascod, top = None ):
    """
    Menghitung frekuensi dan persentase dari data media habit

    Parameter
    ---------
    data_used : data yang digunakan untuk menghitung
    kode_awal : awalan nama kolom dari data yang digunakan
    top : batas berapa data teratas yang akan dihitung

    Return
    ------
    tab_presn_tamb : tabel luaran hasil
    """

    # Filtering kolom data
    data_tambahan = filter_tambahan(data_used, kode_awal, incl_kota = False)

    # Menggabungkan data ke 1 kolom
    data_tamb_merged = pd.Series(data_tambahan.values.ravel('F'))
    data_tamb_merged = data_tamb_merged.to_frame()

    # Labeling data
    for i in range(len(mascod)):
        coding_i = mascod.iloc[i,0]
        data_tamb_merged.loc[data_tamb_merged.iloc[:,0]==coding_i,0] = mascod.iloc[i,1]

    # Membuat tabel count
    tab_count = data_tamb_merged.value_counts().to_frame()
    tab_count.index.rename('Kriteria', inplace = True)
    tab_count.rename(columns = {0 : 'Count'}, inplace = True)

    # Menghitung n_sampel
    kol_1 = data_tambahan.iloc[:,0].to_frame()
    kol_1 = kol_1.dropna()
    n_sampel = len(kol_1)

    # Menghitung presentase
    tab_presn_tamb = pd.DataFrame(index = tab_count.index, columns = ['Count', 'Presentase'])
    tab_presn_tamb['Count'] = tab_count['Count']
    tab_presn_tamb['Presentase'] = (tab_count['Count']/n_sampel)*100

    # Menghitung total multirespon
    tot_multirespon = sum(tab_presn_tamb['Presentase'])

    # Pembulatan
    tab_presn_tamb = round(tab_presn_tamb, 1)

    # Slicing data (Mengambil beberapa teratas)
    if top != None:
        tab_presn_tamb = tab_presn_tamb.iloc[0:top, :]

    # Menambahkan n sampel dan multirespon ke baris terakhir
    row_n_multi = pd.DataFrame({'Kriteria' : ['Multirespon'], 'Count' : [n_sampel], 'Presentase' : [tot_multirespon]})
    tab_presn_tamb.reset_index(inplace = True)
    tab_presn_tamb = pd.concat([tab_presn_tamb, row_n_multi], axis = 0)
    tab_presn_tamb.set_index('Kriteria', inplace = True)

    # Pembulatan
    tab_presn_tamb = round(tab_presn_tamb, 1)

    return tab_presn_tamb

def data_olshop(data, mascod_kota, mascod_olshop):
    """
    Membuat tabel data perbandingan n Sampel dengan Olshop 'Pernah' dan 'Tidak pernah'

    Parameter
    ---------
    data : data yang akan diolah

    Return
    ------
    tabel_totalBawah : tabel yang berisi jumlah n sampel dalam format angka biasa
    persen_tabel : tabel yang berisi jumlah n sampel dalam format persentase
    """

    # Labeling data
    data.columns = ['kota','Olshop']

    for i in range(len(mascod_kota)):
        coding_i = mascod_kota.loc[i,'Coding']
        data.loc[data.loc[:,'kota']==coding_i,'kota'] = mascod_kota.loc[i,'Label']

    for i in range(len(mascod_olshop)):
        coding_i = mascod_olshop.loc[i,'Coding']
        data.loc[data.loc[:,'Olshop']==coding_i,'Olshop'] = mascod_olshop.loc[i,'Label']


    crosstab = pd.crosstab(data['Olshop'],
                              data['kota'])
    crosstab = crosstab.reindex(columns = kota_urut())

    total_bawah = crosstab.sum(axis=0).to_frame().transpose()
    total_bawah.index = ['Total']
    total_bawah.index.rename('Olshop',inplace=True)

    tabel_totalBawah = pd.concat([crosstab,
                            total_bawah],
                           axis=0)
    tabel_totalBawah['Total'] = tabel_totalBawah.sum(axis=1)
    tabel_totalBawah = tabel_totalBawah.transpose()
    tabel_totalBawah.reset_index(inplace=True)
    tabel_totalBawah.columns.name=None
    tabel_totalBawah.set_index('kota',inplace=True)
    tabel_totalBawah.reset_index(inplace=True)

    persen_tabel = tabel_totalBawah.copy()
    persen_tabel['Tidak pernah'] = (persen_tabel['Tidak pernah']/persen_tabel.Total)*100
    persen_tabel['Tidak pernah'] = round(persen_tabel['Tidak pernah'], 1)
    persen_tabel['Pernah'] = (persen_tabel['Pernah']/persen_tabel.Total)*100
    persen_tabel['Pernah'] = round(persen_tabel['Pernah'], 1)
    persen_tabel.reset_index(drop=True, inplace=True)

    return tabel_totalBawah, persen_tabel
