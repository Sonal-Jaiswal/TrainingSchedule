import * as React from 'react';
import type { ITrainingCardProps } from './ITrainingCardProps';
const TrainingCard: React.FC<ITrainingCardProps> = (props) => {
 return (
<div style={{
     backgroundColor: 'white',
     padding: '22px',
     borderRadius: '10px',
     width: '280px',
     boxShadow: '0 2px 8px rgba(0,0,0,0.08)'
   }}>
<div style={{
       fontSize: '12px',
       color: '#2563eb',
       fontWeight: 'bold'
     }}>
       {props.category}
</div>
<h3 style={{
       margin: '10px 0',
       color: '#1f2937'
     }}>
       {props.title}
</h3>
<p style={{
       color: '#6b7280',
       fontSize: '14px'
     }}>
       {props.description}
</p>
<p style={{
       color: '#6b7280',
       fontSize: '14px'
     }}>
       📅 {props.date}
</p>
<p style={{
       color: '#6b7280',
       fontSize: '14px'
     }}>
       👥 {props.seats} seats available
</p>
<button style={{
       width: '100%',
       padding: '10px',
       border: 'none',
       borderRadius: '6px',
       backgroundColor: '#2563eb',
       color: 'white',
       fontWeight: 'bold',
       cursor: 'pointer'
     }}>
       Enroll
</button>
</div>
 );
};
export default TrainingCard;