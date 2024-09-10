import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './AniversarioWebPartWebPart.module.scss';
import * as strings from 'AniversarioWebPartWebPartStrings';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";



export interface IAniversarioWebPartWebPartProps {
  description: string;
}

export type AniversarioItem = {
  Title: string;
  DataNascimento: string;
  Email: string;
  Filial: {
    Title: string;
  };
  Departamento: {
    Title: string;
  };


}



export default class AniversarioWebPartWebPart extends BaseClientSideWebPart<IAniversarioWebPartWebPartProps> {

 
_aniversariantes?: AniversarioItem[];

 

  public render(): void {
    const hoje = new Date();
    const mesAtual = hoje.getMonth();
    const diaHoje = hoje.getDate();


    let aniversariantesMesAtual = this._aniversariantes?.filter(item => {
      const dataNascimento = new Date(item.DataNascimento);
      return dataNascimento.getMonth() === mesAtual && dataNascimento.getDate() >= diaHoje;
    }).sort((a, b) =>new Date(a.DataNascimento).getDate() - new Date(b.DataNascimento).getDate()).slice(0, 4);

    if(!aniversariantesMesAtual || aniversariantesMesAtual.length < 4 ){
      const proxMes = (mesAtual + 1) % 12;
      const aniversariantesMesProximo = this._aniversariantes?.filter(item => {
        const dataNascimento = new Date(item.DataNascimento); 
        return dataNascimento.getMonth() === proxMes;
        }).sort((a, b) => new Date(a.DataNascimento).getDate() - new Date(b.DataNascimento).getDate()).slice(0, 4 - (aniversariantesMesAtual?.length || 0));

        aniversariantesMesAtual = [...(aniversariantesMesAtual || []), ...(aniversariantesMesProximo || [])];

    }

    this.domElement.innerHTML = `<section>
             <div>
                  <div style="padding-bottom: 0.5rem; font-size: 1.5em">
                  <strong>Aniversariantes</strong>          
            </div>           
            <div style="max-height: 60vh">            
    ${aniversariantesMesAtual?.map(item =>
      ` <div class="${styles.cardContainer}">
        <div class="${styles.cardDate}"> 
        <div class="${styles.centralize}" style="height: 50%">
        ${new Date(item.DataNascimento).toLocaleString('default', { month:'short'}).toUpperCase()}
    . </div> 
  <div class="${styles.centralize}" style="height: 50%; font-size: 1.5em"> 
<strong>
${new Date(item.DataNascimento).getDate()}
</strong> 
</div>
 </div> 
 <div style="padding-left: 1rem; font-size: 1em; font-weight: 400">
  <div style="padding: 0.25rem">
   <strong>
    ${item.Title}
   </strong> 
</div> 
<div style="padding: 0.25rem">
${item.Departamento?.Title} - ${item.Filial?.Title}
</div> 
<div style="padding: 0.25rem; font-size: 0.9em;">
 <a href="mailto:${item.Email}">
${item.Email}
</a> 
</div>
</div> 
</div> `
).join(
''
)}
</div> 
</div> </section>
`;
  }

  protected async onInit(): Promise<void> {
    
    return super.onInit().then(async _ => { 
      this._aniversariantes = await this.carregarListaAniversario();
      
    });


  }
  private async carregarListaAniversario(): Promise<AniversarioItem[]>{
    const sp = spfi().using(SPFx(this.context));
    const response: AniversarioItem[] = await sp.web.lists.getByTitle("Colaboradores").items.select("Title", "Email", "DataNascimento", "Filial","Filial/Title", "Departamento" , "Departamento/Title").expand("Filial","Departamento").top(300)();
    console.log(response)
    return response;
  }


  

 

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
