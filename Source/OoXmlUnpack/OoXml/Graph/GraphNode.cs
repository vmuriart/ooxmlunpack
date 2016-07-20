namespace OoXml.Graph
{
    using System.Collections.Generic;
    using System.Diagnostics;

    [DebuggerDisplay("{DebuggerDisplay(),nq}")]
    internal class GraphNode<TKey, TValue>
    {
        public GraphNode(TKey key, TValue value)
        {
            this.Key = key;
            this.Value = value;
            this.Precedents = new Dictionary<TKey, GraphNode<TKey, TValue>>();
            this.Dependents = new Dictionary<TKey, GraphNode<TKey, TValue>>();
        }

        public TKey Key { get; private set; }

        public TValue Value { get; private set; }

        public Dictionary<TKey, GraphNode<TKey, TValue>> Precedents { get; private set; }

        public Dictionary<TKey, GraphNode<TKey, TValue>> Dependents { get; private set; }

        public void DependsOn(GraphNode<TKey, TValue> node)
        {
            this.Precedents[node.Key] = node;
            node.Dependents[this.Key] = this;
        }

        private string DebuggerDisplay()
        {
            return string.Format("{0} : {1}", this.Key, this.Value);
        }
    }
}