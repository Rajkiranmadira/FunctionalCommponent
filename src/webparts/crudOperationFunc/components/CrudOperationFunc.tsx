// import * as React from 'react';
// import styles from './CrudOperationFunc.module.scss';
// import { ICrudOperationFuncProps } from './ICrudOperationFuncProps';
// import { escape } from '@microsoft/sp-lodash-subset';

// export default class CrudOperationFunc extends React.Component<ICrudOperationFuncProps, {}> {
//   public render(): React.ReactElement<ICrudOperationFuncProps> {
//     const {
//       description,
//       isDarkTheme,
//       environmentMessage,
//       hasTeamsContext,
//       userDisplayName
//     } = this.props;

//     return (
//       <section className={`${styles.crudOperationFunc} ${hasTeamsContext ? styles.teams : ''}`}>
//         <div className={styles.welcome}>
//           <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
//           <h2>Well done, {escape(userDisplayName)}!</h2>
//           <div>{environmentMessage}</div>
//           <div>Web part property value: <strong>{escape(description)}</strong></div>
//         </div>
//         <div>
//           <h3>Welcome to SharePoint Framework!</h3>
//           <p>
//             The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
//           </p>
//           <h4>Learn more about SPFx development:</h4>
//           <ul className={styles.links}>
//             <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
//             <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
//           </ul>
//         </div>
//       </section>
//     );
//   }
// }
import * as React from 'react';
import { ICrudOperationFuncProps } from './ICrudOperationFuncProps';
import {useState} from 'react';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.min.js';
import {escape} from '@microsoft/sp-lodash-subset';
import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';

interface EmployeeDetailsStates{
  ID:number;
  Title:string;
  Age:number;
}
const CrudOperationFunc:React.FC<ICrudOperationFuncProps>=(props:ICrudOperationFuncProps)=>{
const [fullName,setFullName]=useState('');
const[age,setAge]=useState('');
const [allItems,setAllItems]=useState<EmployeeDetailsStates[]>([]);

const createItem=async():Promise<void>=>{

  const body:string=JSON.stringify({
    'Title':fullName,
    'Age':age
  });
  try{
    const response:SPHttpClientResponse=await props.context.spHttpClient.post(
      `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeDetails')/items`,
      SPHttpClient.configurations.v1,
      {
        headers:{
          'Accept':'application/json;odata=nometadata',
          'Content-type':'application/json;odata=nometadata',
          'odata-version':''
        },
        body:body
      }
    );
    if(response.ok){
      const responseJSON=await response.json();
      console.log(responseJSON);
      alert(`Item created successfully with ID : ${responseJSON.ID}`);
    }
    else{
      const responseJSON=await response.json();
      console.log(responseJSON);
      alert(`Something went wrong! Check the error in the browser console.`);
    }

  }
  catch(error){
    console.log(error);
    alert(`An error occurred while creating the item`);

  }
}
const getItemById=():void=>{
  const idElement=document.getElementById('itemId') as HTMLInputElement |null;
  if(idElement?.value){
    const id:number=Number(idElement.value);
    if(id>0){
      props.context.spHttpClient.get(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeDetails')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers:{
          'Accept':'application/json;odata=nometadata',
          'odata-version':''
        }
      }
      )
      .then((response:SPHttpClientResponse)=>{
        if(response.ok){
          response.json().then((responseJSON)=>{
            console.log(responseJSON);
            setFullName(responseJSON.Title);
            setAge(responseJSON.Age);
          });
        }
        else{
          response.json().then((responseJSON)=>{
            console.log(responseJSON);
            alert(`Something went wrong ! check the error in the browser`);
          });
        }
      })
      .catch((error:any)=>{
        console.log(error);
      });
    }
    else{
      alert(`Please enter a valid item id`);
    }
  }
  else{
    console.log("Error: Element 'itemId' not found ");
  }
};

const getAllItem=():void=>{

  props.context.spHttpClient.get(props.context.pageContext.web.absoluteUrl+ `/_api/web/lists/getbytitle('EmployeeDetails')/items`,
  SPHttpClient.configurations.v1,
  {
    headers:{
      'Accept':'application/json;odata=nometadata',
      'odata-version':''
    }
  }
  )
  .then((response:SPHttpClientResponse)=>{
    if(response.ok){
      response.json().then((responseJSON)=>{
        setAllItems(responseJSON.value);
        console.log(responseJSON);
      });
    }
    else{
      response.json().then((responseJSON)=>{
        console.log(responseJSON);
        alert(`Something went wrong ! check the error in the browser console `);
      });
    }
  })
  .catch((error:any)=>{
    console.log(error);
  })
}

const updateItem=():void=>{
  const idElement=document.getElementById('itemId') as HTMLInputElement;

  if(idElement){
    const id:number=parseInt(idElement.value);
    const body:string=JSON.stringify({
      'Title':fullName,
      'Age':parseInt(age)
    });
    if(id>0){
      props.context.spHttpClient.post(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeDetails')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers:{
          'Accept':'application/json;odata=nometadata',
          'Content-type':'application/json;odata=nometadata',
          'odata-version':'',
          'IF-MATCH':'*',
          'X-HTTP-Method':'MERGE',
        },
        body:body
      }
      )
      .then((response:SPHttpClientResponse)=>{
        if(response.ok){
          alert(`Item with ID :${id} updated successfully `);
        }
        else{
          response.json().then((responseJSON)=>{
            console.log(responseJSON);
            alert(`Something went wrong ! check the console for finding`);
          });
        }
      })
      .catch((error)=>{
        console.log(error);
      });
    }
    else{
      alert('Please enter the valid id');
    }
  }
  else{
    console.log('Id element is not found');
  }
}
const deleteItem=():void=>{
  const idElement=document.getElementById('itemId') as HTMLInputElement;
  const id:number=parseInt(idElement?.value||'0');
  if(id>0){
    props.context.spHttpClient.post(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeDetails')/items(${id})`,
   SPHttpClient.configurations.v1,
   {
    headers:{
      'Accept':'application/json;odata=nometadata',
      'Content-type':'application/json;odata=verbose',
      'odata-version':'',
      'IF-MATCH':'*',
      'X-HTTP-Method':'DELETE',
    }
   }
   
    )
    .then((response:SPHttpClientResponse)=>{
      if(response.ok){
        alert(`Item Id : ${id} deleted successfully`)
      }
      else{
        alert('Something went wrong');
        console.log(response.json());
      }
    });
  }
  else{
    alert('Please enter valid Id to delete Item');
  }
}
  return (
    <>
    <div className="container">
<div className="row">
<div className="col-md-6">
<p>{escape(props.description)}</p>
<div className="form-group">
<label htmlFor="itemId">Item ID:</label>
<input type="text" className="form-control" id="itemId"></input>
</div>
<div className="form-group">
<label htmlFor="fullName">Full Name</label>
<input type="text" className="form-control" id="fullName" value={fullName} onChange={(e) => setFullName(e.target.value)}></input>
</div>
<div className="form-group">
<label htmlFor="age">Age</label>
<input type="text" className="form-control" id="age" value={age} onChange={(e) => setAge(e.target.value)}></input>
</div>
<div className="form-group">
<label htmlFor="allItems">All Items:</label>
<div id="allItems">
<table className="table table-bordered">
<thead>
<tr>
<th>ID</th>
<th>Full Name</th>
<th>Age</th>
</tr>
</thead>
<tbody>
                    {allItems.map((item) => (
<tr key={item.ID}>
<td>{item.ID}</td>
<td>{item.Title}</td>
<td>{item.Age}</td>
</tr>
                    ))}
</tbody>
</table>
</div>
</div>
<div className="d-flex justify-content-start">
<button className="btn btn-primary mx-2" onClick={createItem}>Create</button>
<button className="btn btn-success mx-2" onClick={getItemById}>Read</button>
<button className="btn btn-info mx-2" onClick={getAllItem}>Read All</button>
<button className="btn btn-warning mx-2" onClick={updateItem}>Update</button>
<button className="btn btn-danger mx-2" onClick={deleteItem}>Delete</button>
</div>
</div>
</div>
</div>

    
    </>
  )

}
export default CrudOperationFunc