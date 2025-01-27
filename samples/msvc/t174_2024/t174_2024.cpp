#include <windows.h>
#include <assert.h>
#include <stdio.h>

#import "C:\Program Files (x86)\Common Files\Testo Shared\tcddka.dll" no_namespace
#import "C:\Program Files (x86)\Common Files\Testo Shared\t17ba.dll" no_namespace

int main(int argc, char** argv)
{

    short nPort(8);

    if (argc > 1)
    {
        nPort = atoi(argv[1]);
        printf("trying to connect using COM %i\n", nPort);
    }
    else
    {
        printf("trying to connect using COM %i (default)\n", nPort);
    }

    OleInitialize(NULL);

    try
    {
        
        ItcddkPtr pTcDdk;
        HRESULT hr = pTcDdk.CreateInstance(__uuidof(tcddk));

        if (S_OK != hr)
        {
            assert(FALSE);
            _com_issue_error(hr);
        }

        VARIANT_BOOL bRet = pTcDdk->Init(nPort -1, "testo174-2024", 7000);

        if (VARIANT_TRUE != bRet)
        {
            assert(FALSE);
            _com_issue_error(E_FAIL);
        }


        //It17caimpPtr pImp = pTcDdk->Instrument;
        IT17baInstrumentPtr pImp = pTcDdk->Instrument;
        if (FALSE == pImp)
        {
            assert(FALSE);
            _com_issue_error(E_NOINTERFACE);
        }

        printf("hello 174\n");
#if 0
        IT17caSocketCollectionPtr psks = pImp->Sockets;
        if (FALSE == psks)
        {
            assert(FALSE);
            _com_issue_error(E_NOINTERFACE);
        }

        IT17caSocketPtr pskt = psks->Item[0];
        if (FALSE == pskt)
        {
            assert(FALSE);
            _com_issue_error(E_NOINTERFACE);
        }

        WCHAR szName[] = L"Pt 100";

        double dMin = pskt->GetMinRange(szName);
        double dMax = pskt->GetMaxRange(szName);

        printf("range is %lf..%lf\n", dMin, dMax);
#endif
#if 0
        // 800. added timestamps

        long lnts = pImp->NumTimeStamps;
        int i = 0;
        for (i = 0; i < lnts; i++)
        {
            long ats;
            if (S_OK == pImp->get_TimeStamp(i, &ats))
            {
                printf("ts %d %ld\n", i, ats);

                // if (1 == i)
                {
                    pImp->put_ActiveTimeStamp(ats);

                    IProtocolCollectionPtr pColl = pTcDdk->Protocols;
                    IProtocolPtr pProt = pColl->Item(0);
                    if (NULL != pProt)
                    {
                        pProt->Get();

                        Itcvi2filePtr q;
                        q.CreateInstance(__uuidof(tcvi2file));
                        if (NULL != q)
                        {
                            q->Add(pProt);

                            q->Save(_bstr_t("c:\\temp\\t17c1109.vi2"));
                        }
                    }

                    pImp->put_ActiveTimeStamp(0);
                }
            }
        }
#endif
    }
    catch (_com_error& e)
    {
        printf("com error %08lx\n", e.Error());
        if (NULL != (BSTR)e.Description())
        {
            printf("description :%S:\n", (BSTR)e.Description());
        }
    }
    catch (...)
    {
        printf("unexpected\n");
    }

    OleUninitialize();

    return 0;
}
