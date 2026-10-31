using System.Collections;
using UnityEngine;

public class DisableTriggerAfterUse : MonoBehaviour
{
    [SerializeField] private string activatorTag = "Player";
    [SerializeField] private Collider triggerCollider;
    [SerializeField] private MonoBehaviour[] behavioursToDisable;

    private bool hasTriggered;

    private void Reset()
    {
        triggerCollider = GetComponent<Collider>();
    }

    private void Awake()
    {
        if (triggerCollider == null)
            triggerCollider = GetComponent<Collider>();
    }

    private void OnTriggerEnter(Collider other)
    {
        if (hasTriggered || !other.CompareTag(activatorTag))
            return;

        hasTriggered = true;
        StartCoroutine(DisableAfterTrigger());
    }

    private IEnumerator DisableAfterTrigger()
    {
        // 같은 진입 이벤트에서 active_objects가 먼저 실행될 시간을 보장한다.
        yield return null;

        foreach (MonoBehaviour behaviour in behavioursToDisable)
        {
            if (behaviour != null)
                behaviour.enabled = false;
        }

        if (triggerCollider != null)
            triggerCollider.enabled = false;

        enabled = false;
    }
}
